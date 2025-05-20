using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using NLog;
using NLog.Config;
using NLog.Targets;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using FormsFlowDirection = System.Windows.Forms.FlowDirection;


// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Drill Namer")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("Drill Namer")]
[assembly: AssemblyCopyright("Copyright ©  2024")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("22473b76-d2f6-4d29-8e92-46d5d4f1134a")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace Drill_Namer
{
    public static class AutoCADHelper
    {
        /// <summary>
        /// Start a transaction in the current AutoCAD document.
        /// </summary>
        public static Transaction StartTransaction()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            return doc.Database.TransactionManager.StartTransaction();
        }

        /// <summary>
        /// Prompts the user to select an insertion point. Returns true if a point was chosen.
        /// </summary>
        /// <param name="promptMessage">Message to display.</param>
        /// <param name="insertionPoint">The point selected by the user.</param>
        /// <returns>True if user picked a point; false if canceled.</returns>
        public static bool GetInsertionPoint(string promptMessage, out Point3d insertionPoint)
        {
            insertionPoint = Point3d.Origin;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return false;

            Editor ed = doc.Editor;
            PromptPointOptions opts = new PromptPointOptions("\n" + promptMessage);
            PromptPointResult res = ed.GetPoint(opts);

            if (res.Status == PromptStatus.OK)
            {
                insertionPoint = res.Value;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Insert a block reference at a specified point, with given attributes and scale,
        /// on layer "CG-NOTES". If missing, tries to load the block from 
        /// C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties\[blockName].dwg
        /// </summary>
        public static ObjectId InsertBlock(
            string blockName,
            Point3d insertionPoint,
            Dictionary<string, string> attributes,
            double scale = 1.0)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId brId = ObjectId.Null;

            // Make sure we lock the doc for changes
            {
                {
                    // 1) Ensure the block is loaded into this DWG
                    if (!EnsureBlockIsLoaded(db, blockName))
                    {
                        // Could not find or load it from the external .dwg
                        return ObjectId.Null;
                    }

                    // 2) Get the block table record for that block
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (!bt.Has(blockName))
                    {
                        // Even after loading, not found
                        return ObjectId.Null;
                    }

                    ObjectId btrId = bt[blockName];
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    // 3) Open ModelSpace (or whichever space you need)
                    //    For standard insertion, usually model space:
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 4) Create the block reference
                    BlockReference br = new BlockReference(insertionPoint, btrId);

                    // Apply the scale
                    br.ScaleFactors = new Scale3d(scale, scale, scale);

                    // Set layer to "CG-NOTES" 
                    // (make sure the layer exists; if not, create it or set to another fallback)
                    if (!CreateLayerIfMissing(db, "CG-NOTES"))
                    {
                        // If for any reason we fail to create/find the layer, fallback to "0"
                        br.Layer = "0";
                    }
                    else
                    {
                        br.Layer = "CG-NOTES";
                    }

                    // 5) Append the block reference to model space
                    ms.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    // 6) Add the attribute references
                    foreach (ObjectId id in btr)
                    {
                        DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                        if (obj is AttributeDefinition attDef)
                        {
                            // Create an AttributeReference from the def
                            AttributeReference attRef = new AttributeReference();
                            attRef.SetAttributeFromBlock(attDef, br.BlockTransform);

                            // If user provided an attribute matching this Tag
                            if (attributes != null && attributes.ContainsKey(attDef.Tag))
                            {
                                attRef.TextString = attributes[attDef.Tag];
                            }
                            else
                            {
                                // If not provided, just keep the attDef's default
                                attRef.TextString = attDef.TextString;
                            }

                            // set the same layer as the block reference
                            attRef.Layer = br.Layer;

                            // Add it to the blockRef
                            br.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }

                    tr.Commit();
                    brId = br.ObjectId;
                }
            }

            return brId;
        }

        /// <summary>
        /// Attempt to ensure the block definition is present in the current DB.
        /// If missing, tries to load it from "C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties\blockName.dwg".
        /// </summary>
        private static bool EnsureBlockIsLoaded(Database db, string blockName)
        {
            bool exists = false;
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                exists = bt.Has(blockName);
                tr.Commit();
            }
            if (exists) return true;

            // If not found, try to import from external .dwg
            string baseFolder = @"C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties";
            string dwgPath = Path.Combine(baseFolder, blockName + ".dwg");
            if (!File.Exists(dwgPath))
            {
                // can't find the external block .dwg
                return false;
            }

            return (ImportBlockDefinition(db, dwgPath, blockName) != ObjectId.Null);
        }

        /// <summary>
        /// Imports a DWG file (which is effectively a single block definition) into the current DB
        /// with DuplicateRecordCloning.Replace (so we re-use the name).
        /// </summary>
        private static ObjectId ImportBlockDefinition(Database destDb, string sourceDwgPath, string blockName)
        {
            if (destDb == null || string.IsNullOrEmpty(sourceDwgPath) || !File.Exists(sourceDwgPath))
                return ObjectId.Null;

            ObjectId result = ObjectId.Null;
            {
                try
                {
                    // read the .dwg containing the block
                    tempDb.ReadDwgFile(sourceDwgPath, FileShare.Read, true, "");
                    {
                        BlockTable sourceBT = (BlockTable)tr.GetObject(tempDb.BlockTableId, OpenMode.ForRead);
                        if (!sourceBT.Has(blockName))
                        {
                            // blockName does not exist in that DWG
                            return ObjectId.Null;
                        }

                        // get the block's btr
                        ObjectId sourceBtrId = sourceBT[blockName];
                        ObjectIdCollection idsToClone = new ObjectIdCollection();
                        idsToClone.Add(sourceBtrId);

                        tr.Commit();

                        // now clone it into destDb
                        IdMapping mapping = new IdMapping();
                        destDb.WblockCloneObjects(idsToClone, destDb.BlockTableId, mapping,
                            DuplicateRecordCloning.Replace, false);
                    }

                    // verify we have it
                    {
                        BlockTable destBT = (BlockTable)destTr.GetObject(destDb.BlockTableId, OpenMode.ForRead);
                        if (destBT.Has(blockName))
                        {
                            result = destBT[blockName];
                        }
                        destTr.Commit();
                    }
                }
                catch
                {
                    // do nothing
                }
            }

            return result;
        }

        /// <summary>
        /// Creates or verifies that layerName exists. Returns true if found or created successfully.
        /// </summary>
        private static bool CreateLayerIfMissing(Database db, string layerName)
        {
            bool success = false;
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt.Has(layerName))
                {
                    success = true;
                }
                else
                {
                    // create the layer
                    lt.UpgradeOpen();
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;
                    // optionally set color, line type, etc.
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);

                    success = true;
                }
                tr.Commit();
            }
            return success;
        }
    }
}

namespace Drill_Namer
{
    public static class Logger
    {
        private static NLog.Logger logger;

        static Logger()
        {
            ConfigureLogger();
        }

        /// <summary>
        /// Configures the NLog logger with a dynamic log file path based on the current drawing.
        /// </summary>
        private static void ConfigureLogger()
        {
            try
            {
                // Create a new NLog configuration
                var config = new LoggingConfiguration();

                // Get the current drawing's directory for the log file path
                string logFilePath = GetLogFilePath();

                // Create a FileTarget with the dynamic log file path
                var logfile = new FileTarget("logfile")
                {
                    FileName = logFilePath,
                    Layout = "${longdate} | ${level:uppercase=true} | ${message}"
                };

                // Add the rule for mapping loggers to the FileTarget
                config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);

                // Apply the configuration
                LogManager.Configuration = config;

                logger = LogManager.GetCurrentClassLogger();
            }
            catch (Exception ex)
            {
                try
                {
                    string source = "DrillNamer";
                    string log = "Application";
                    if (!EventLog.SourceExists(source))
                    {
                        EventLog.CreateEventSource(source, log);
                    }
                    EventLog.WriteEntry(source, $"Error initializing NLog: {ex.Message}", EventLogEntryType.Error);
                }
                catch { }
            }
        }

        /// <summary>
        /// Retrieves the log file path based on the current drawing's directory.
        /// If the drawing is unsaved, defaults to the user's Documents folder.
        /// </summary>
        /// <returns>Full path to the log file.</returns>
        private static string GetLogFilePath()
        {
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                if (acDoc != null && acDoc.IsNamedDrawing && !string.IsNullOrEmpty(acDoc.Name))
                {
                    string drawingDirectory = Path.GetDirectoryName(acDoc.Name);
                    string logFilePath = Path.Combine(drawingDirectory, "DrillNamer.log");
                    return logFilePath;
                }
                else
                {
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "DrillNamer.log");
                }
            }
            catch
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "DrillNamer.log");
            }
        }

        /// <summary>
        /// Updates the log file path based on the current drawing's directory.
        /// This method is called before each logging action to ensure the log file is correctly located.
        /// </summary>
        private static void UpdateLogFilePath()
        {
            try
            {
                string logFilePath = GetLogFilePath();
                var config = LogManager.Configuration;
                var logfile = config.FindTargetByName<FileTarget>("logfile");
                if (logfile != null)
                {
                    // Use Render to obtain the actual file name
                    string currentFileName = logfile.FileName.Render(new LogEventInfo());
                    if (!string.Equals(currentFileName, logFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        logfile.FileName = logFilePath;
                        LogManager.ReconfigExistingLoggers();
                    }
                }
            }
            catch
            {
                // Silently fail
            }
        }

        public static void LogInfo(string message)
        {
            UpdateLogFilePath();
            logger.Info(message);
        }

        public static void LogWarning(string message)
        {
            UpdateLogFilePath();
            logger.Warn(message);
        }

        public static void LogError(string message)
        {
            UpdateLogFilePath();
            logger.Error(message);
        }

        public static void LogDebug(string message)
        {
            UpdateLogFilePath();
            logger.Debug(message);
        }
    }
}

namespace Drill_Namer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Enable visual styles and set text rendering
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Run the main form
            Application.Run(new FindReplaceForm());
        }
    }
}

namespace Drill_Namer
{
    public class FindReplaceCommand
    {
        // The CommandMethod attribute makes this function callable from the AutoCAD command line.
        [CommandMethod("drillnames")]
        public static void ShowFindReplaceForm()
        {
            // Create and show the form
            FindReplaceForm form = new FindReplaceForm();
            // Shows the form as a modal dialog (i.e., waits for user input before continuing).
            acadApp.ShowModalDialog(form);
        }
    }
}
namespace Drill_Namer
{
    partial class FindReplaceForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        // This function initializes the form components (labels, textboxes, buttons)
        private void InitializeComponent()
        {
            // SuspendLayout pauses the layout of controls until ResumeLayout is called, improving performance.
            this.SuspendLayout();

            // Define the form size
            this.ClientSize = new System.Drawing.Size(600, 650); // Adjust size as needed
            this.Name = "FindReplaceForm"; // Name of the form
            this.Text = "Find and Replace Form"; // Title of the form window

            // Add other controls like labels, textboxes, and buttons dynamically
            // This will be handled in the main .cs file for flexibility

            // Resumes the layout of controls
            this.ResumeLayout(false);
        }

        #endregion
    }
}

// Define alias to resolve ambiguity between Application classes

// Define alias for System.Windows.Forms.FlowDirection

namespace Drill_Namer
{
    public partial class FindReplaceForm : Form
    {
        private TextBox[] drillTextBoxes;
        private Label[] drillLabels;
        private const int DrillCount = 12; // Number of drills

        public FindReplaceForm()
        {
            InitializeComponent();
            InitializeDynamicControls();
            LoadFromJson(); // Load from JSON when the form initializes
        }

        /// <summary>
        /// Generates the JSON file path based on the drawing name.
        /// </summary>
        /// <returns>Full path to the JSON file.</returns>
        private string GetJsonFilePath()
        {
            var acDoc = AcApplication.DocumentManager.MdiActiveDocument;
            if (acDoc == null || string.IsNullOrEmpty(acDoc.Name))
            {
                MessageBox.Show("No active AutoCAD document found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new InvalidOperationException("No active AutoCAD document found.");
            }

            string dwgDirectory = Path.GetDirectoryName(acDoc.Name);  // DWG file's directory
            string drawingName = Path.GetFileNameWithoutExtension(acDoc.Name);  // DWG file's name without extension

            // Remove any trailing hyphens to prevent empty segments
            drawingName = drawingName.TrimEnd('-');

            // Split the drawing name by '-' and take the first two segments for the prefix
            string[] nameParts = drawingName.Split('-');
            string prefix = nameParts.Length >= 2 ? $"{nameParts[0]}-{nameParts[1]}" : drawingName;

            // Construct the JSON file path
            string jsonFilePath = Path.Combine(dwgDirectory, $"{prefix}.json");

            return jsonFilePath;
        }

        /// <summary>
        /// Initializes all dynamic controls (labels, textboxes, buttons).
        /// </summary>
        private void InitializeDynamicControls()
        {
            drillTextBoxes = new TextBox[DrillCount];
            drillLabels = new Label[DrillCount];

            int controlHeight = 25;  // Standard height for individual buttons
            int labelX = 10;         // X position for labels
            int textboxX = 170;      // X position for textboxes
            int buttonSetX = 330;    // X position for SET button
            int buttonCreateX = 410; // X position for CREATE button
            int buttonResetX = 490;  // X position for RESET button
            int verticalSpacing = 5; // Spacing between rows

            for (int i = 0; i < DrillCount; i++)
            {
                int currentIndex = i; // Local copy to capture the correct index
                int controlY = 20 + currentIndex * (controlHeight + verticalSpacing); // Y position for each row

                // Label for each drill
                drillLabels[currentIndex] = new Label
                {
                    Text = $"DRILL_{currentIndex + 1}",
                    Location = new System.Drawing.Point(labelX, controlY),
                    Size = new System.Drawing.Size(150, controlHeight)
                };
                this.Controls.Add(drillLabels[currentIndex]);

                // TextBox for each drill
                drillTextBoxes[currentIndex] = new TextBox
                {
                    Location = new System.Drawing.Point(textboxX, controlY),
                    Size = new System.Drawing.Size(150, controlHeight)
                };
                this.Controls.Add(drillTextBoxes[currentIndex]);

                // SET Button
                var setButton = new Button
                {
                    Text = "SET",
                    Location = new System.Drawing.Point(buttonSetX, controlY),
                    Size = new System.Drawing.Size(75, controlHeight)
                };
                setButton.Click += (sender, e) => SetDrill(currentIndex);
                this.Controls.Add(setButton);

                // CREATE Button
                var createButton = new Button
                {
                    Text = "CREATE",
                    Location = new System.Drawing.Point(buttonCreateX, controlY),
                    Size = new System.Drawing.Size(75, controlHeight)
                };
                createButton.Click += (sender, e) => CreateButton_Click(currentIndex);
                this.Controls.Add(createButton);

                // RESET Button
                var resetButton = new Button
                {
                    Text = "RESET",
                    Location = new System.Drawing.Point(buttonResetX, controlY),
                    Size = new System.Drawing.Size(75, controlHeight)
                };
                resetButton.Click += (sender, e) => ResetDrill(currentIndex);
                this.Controls.Add(resetButton);
            }

            // Adjusted heights for bottom section buttons
            var buttonHeight = 30;  // Slightly larger for "ALL" buttons
            var buttonWidth = 150;   // Increased width for better visibility
            var bottomButtonY = DrillCount * (controlHeight + verticalSpacing) + 20;

            // Set All Button
            var setAllButton = new Button
            {
                Text = "SET ALL",
                Size = new System.Drawing.Size(buttonWidth, buttonHeight)
            };
            setAllButton.Click += SetAllButton_Click;

            // Create All Button
            var createAllButton = new Button
            {
                Text = "CREATE ALL",
                Size = new System.Drawing.Size(buttonWidth, buttonHeight)
            };
            createAllButton.Click += CreateAllButton_Click;

            // Reset All Button
            var resetAllButton = new Button
            {
                Text = "RESET ALL",
                Size = new System.Drawing.Size(buttonWidth, buttonHeight)
            };
            resetAllButton.Click += ResetAllButton_Click;

            // Update From Block Attribute Button
            var updateButton = new Button
            {
                Text = "UPDATE FROM BLOCK ATTRIBUTE",
                Size = new System.Drawing.Size(220, buttonHeight) // Increased width to fit text
            };
            updateButton.Click += UpdateFromAttributesButton_Click;

            // Create a FlowLayoutPanel for "All" buttons to manage layout
            FlowLayoutPanel allButtonsPanel = new FlowLayoutPanel
            {
                Location = new System.Drawing.Point(buttonSetX, bottomButtonY),
                Size = new System.Drawing.Size(800, buttonHeight + 10), // Adjust size as needed
                FlowDirection = FormsFlowDirection.LeftToRight, // Using alias
                WrapContents = false
            };
            this.Controls.Add(allButtonsPanel);

            // Add "All" buttons to the FlowLayoutPanel
            allButtonsPanel.Controls.Add(setAllButton);
            allButtonsPanel.Controls.Add(createAllButton);
            allButtonsPanel.Controls.Add(resetAllButton);
            allButtonsPanel.Controls.Add(updateButton);

            // Adjust the form's size to accommodate all buttons without overlapping
            this.Width = 850;  // Increased width to prevent overlapping
            this.Height = bottomButtonY + buttonHeight + 100; // Adjust height accordingly
        }

        /// <summary>
        /// Sets a specific drill to the user-defined name if it doesn't have the default value.
        /// </summary>
        /// <param name="index">Index of the drill (0-based).</param>
        private void SetDrill(int index)
        {
            if (index < 0 || index >= DrillCount) // Ensure valid index
            {
                MessageBox.Show($"Invalid index: {index}. Must be between 0 and {DrillCount - 1}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string newDrillName = drillTextBoxes[index].Text.Trim();

            if (string.IsNullOrWhiteSpace(newDrillName))
            {
                MessageBox.Show($"Drill name for DRILL_{index + 1} cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string defaultValue = $"DRILL_{index + 1}";

            // Check if current value is not default before setting
            if (newDrillName.Equals(defaultValue, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show($"DRILL_{index + 1} is already at its default value.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Replace text in the drawing
            ReplaceTextInAutoCAD(defaultValue, newDrillName);

            // Replace block attributes in the drawing
            ReplaceBlockAttributesInAutoCAD(index + 1, newDrillName);

            // Update the label to reflect the new value
            drillLabels[index].Text = newDrillName;

            // Save the updated data to JSON
            SaveData();

            MessageBox.Show($"DRILL_{index + 1} has been set to '{newDrillName}'.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Handles the CREATE button click for an individual drill.
        /// </summary>
        /// <param name="index">Index of the drill (0-based).</param>
        private void CreateButton_Click(int index)
        {
            if (index < 0 || index >= DrillCount)
            {
                MessageBox.Show($"Invalid index: {index}. Must be between 0 and {DrillCount - 1}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string drillName = drillTextBoxes[index].Text.Trim();

            if (string.IsNullOrWhiteSpace(drillName))
            {
                MessageBox.Show($"Drill name for DRILL_{index + 1} cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            InsertBlockIntoLayout(drillName, index + 1); // Insert the block into the layout

            // Inform the user that the block was created
            MessageBox.Show($"Block for DRILL_{index + 1} ('{drillName}') has been created.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Resets a specific drill to its default value or blank based on its index.
        /// </summary>
        /// <param name="index">Index of the drill (0-based).</param>
        private void ResetDrill(int index)
        {
            if (index < 0 || index >= DrillCount) // Ensure valid index
            {
                MessageBox.Show($"Invalid index: {index}. Must be between 0 and {DrillCount - 1}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string defaultValue = $"DRILL_{index + 1}";
            string currentValue = drillTextBoxes[index].Text.Trim();

            // Determine if a reset is needed
            bool needsReset = false;

            if (index + 1 == 1)
            {
                // For DRILL_1, always reset to "DRILL_1" if not already
                if (!currentValue.Equals("DRILL_1", StringComparison.OrdinalIgnoreCase))
                {
                    needsReset = true;
                }
            }
            else
            {
                // For DRILL_2 to DRILL_12, reset only if current value is not blank
                if (!string.IsNullOrWhiteSpace(currentValue))
                {
                    needsReset = true;
                }
            }

            if (needsReset)
            {
                // Replace text in the drawing
                if (index + 1 == 1)
                {
                    // DRILL_1 resets to "DRILL_1"
                    ReplaceTextInAutoCAD(currentValue, defaultValue);
                }
                else
                {
                    // DRILL_2-12 resets to blank
                    ReplaceTextInAutoCAD(currentValue, "");
                }

                // Replace block attributes in the drawing
                if (index + 1 == 1)
                {
                    // If DRILL_1, set to "DRILL_1"
                    ReplaceBlockAttributesInAutoCAD(index + 1, "DRILL_1", isReset: true);
                }
                else
                {
                    // For other DRILL_n, set to blank
                    ReplaceBlockAttributesInAutoCAD(index + 1, "", isReset: true);
                }

                // Reset the TextBox and Label
                if (index + 1 == 1)
                {
                    drillTextBoxes[index].Text = defaultValue;
                    drillLabels[index].Text = defaultValue;
                }
                else
                {
                    drillTextBoxes[index].Text = "";
                    drillLabels[index].Text = "";
                }

                // Save the updated data to JSON
                SaveData();

                MessageBox.Show($"DRILL_{index + 1} has been reset.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // No action needed as the drill is already at its default value
                MessageBox.Show($"DRILL_{index + 1} is already at its default value.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Replaces text in AutoCAD drawings (DBText and MText), including within nested blocks.
        /// </summary>
        /// <param name="oldValue">Text to replace.</param>
        /// <param name="newValue">New text value.</param>
        private void ReplaceTextInAutoCAD(string oldValue, string newValue)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            {
                BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForWrite) as BlockTableRecord;

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;

                        if (ent == null || ent.IsErased)
                            continue; // Skip erased or null entities

                        if (ent is DBText dbText && dbText.TextString.Equals(oldValue, StringComparison.OrdinalIgnoreCase))
                        {
                            dbText.TextString = newValue;
                            ed.WriteMessage($"\nReplaced text '{oldValue}' with '{newValue}' in DBText.");
                        }
                        else if (ent is MText mText && mText.Contents.Contains(oldValue))
                        {
                            try
                            {
                                mText.UpgradeOpen();
                                mText.Contents = mText.Contents.Replace(oldValue, newValue);
                                ed.WriteMessage($"\nReplaced text '{oldValue}' with '{newValue}' in MText.");
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                // Handle cases where MText might be locked or otherwise inaccessible
                                ed.WriteMessage($"\nFailed to replace text in MText: {ex.Message}");
                            }
                        }
                        else if (ent is BlockReference blockRef && !blockRef.IsErased)
                        {
                            // Handle nested blocks recursively
                            ReplaceTextInBlock(blockRef, oldValue, newValue, tr, ed);
                        }
                    }
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Recursively replaces text within a BlockReference and its nested blocks.
        /// </summary>
        /// <param name="blockRef">The BlockReference entity.</param>
        /// <param name="oldValue">Text to replace.</param>
        /// <param name="newValue">New text value.</param>
        /// <param name="tr">The current transaction.</param>
        /// <param name="ed">The AutoCAD editor.</param>
        private void ReplaceTextInBlock(BlockReference blockRef, string oldValue, string newValue, Transaction tr, Editor ed)
        {
            if (blockRef == null || blockRef.IsErased)
                return; // Skip erased or null blocks

            // Iterate through each attribute in the block
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                if (attRef != null && attRef.TextString.Contains(oldValue))
                {
                    try
                    {
                        attRef.TextString = attRef.TextString.Replace(oldValue, newValue);
                        ed.WriteMessage($"\nReplaced text '{oldValue}' with '{newValue}' in BlockReference Attribute.");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        // Handle cases where AttributeReference might be locked or inaccessible
                        ed.WriteMessage($"\nFailed to replace text in AttributeReference: {ex.Message}");
                    }
                }
            }

            // Iterate through entities within the block's BlockTableRecord for nested blocks
            BlockTableRecord nestedBtr = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
            foreach (ObjectId objId in nestedBtr)
            {
                Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;

                if (ent == null || ent.IsErased)
                    continue; // Skip erased or null entities

                if (ent is DBText dbText && dbText.TextString.Equals(oldValue, StringComparison.OrdinalIgnoreCase))
                {
                    dbText.TextString = newValue;
                    ed.WriteMessage($"\nReplaced text '{oldValue}' with '{newValue}' in DBText within nested BlockReference.");
                }
                else if (ent is MText mText && mText.Contents.Contains(oldValue))
                {
                    try
                    {
                        mText.UpgradeOpen();
                        mText.Contents = mText.Contents.Replace(oldValue, newValue);
                        ed.WriteMessage($"\nReplaced text '{oldValue}' with '{newValue}' in MText within nested BlockReference.");
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        // Handle cases where MText might be locked or inaccessible
                        ed.WriteMessage($"\nFailed to replace text in nested MText: {ex.Message}");
                    }
                }
                else if (ent is BlockReference nestedBlockRef && !nestedBlockRef.IsErased)
                {
                    // Recursive call for deeper nested blocks
                    ReplaceTextInBlock(nestedBlockRef, oldValue, newValue, tr, ed);
                }
            }
        }

        /// <summary>
        /// Replaces block attribute values in AutoCAD.
        /// </summary>
        /// <param name="drillIndex">Index of the drill (1-based).</param>
        /// <param name="newValue">New drill name value to set.</param>
        /// <param name="isReset">Indicates if the operation is a reset.</param>
        private void ReplaceBlockAttributesInAutoCAD(int drillIndex, string newValue, bool isReset = false)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            {
                BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForWrite) as BlockTableRecord;

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;

                        if (ent == null || ent.IsErased)
                            continue; // Skip erased or null entities

                        if (ent is BlockReference blockRef && !blockRef.IsErased)
                        {
                            // Iterate through each attribute in the block
                            foreach (ObjectId attId in blockRef.AttributeCollection)
                            {
                                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;

                                if (attRef != null)
                                {
                                    // Check if the attribute tag matches DRILL_n where n is the drillIndex
                                    string expectedTag = $"DRILL_{drillIndex}";

                                    if (attRef.Tag.Equals(expectedTag, StringComparison.OrdinalIgnoreCase))
                                    {
                                        try
                                        {
                                            if (isReset && drillIndex == 1)
                                            {
                                                // For DRILL_1 during reset, set to "DRILL_1"
                                                attRef.TextString = "DRILL_1";
                                            }
                                            else if (isReset)
                                            {
                                                // For other DRILL_n during reset, set to blank
                                                attRef.TextString = "";
                                            }
                                            else
                                            {
                                                // For SET operation, set to the new drill name
                                                attRef.TextString = newValue;
                                            }

                                            ed.WriteMessage($"\nUpdated attribute '{attRef.Tag}' to '{attRef.TextString}'.");
                                        }
                                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                        {
                                            // Handle cases where AttributeReference might be locked or inaccessible
                                            ed.WriteMessage($"\nFailed to update AttributeReference: {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Handles the SET ALL button click.
        /// </summary>
        private void SetAllButton_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show("Are you sure you want to set all drills with your inputs?", "Confirm SET ALL", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmResult == DialogResult.Yes)
            {
                bool anySetPerformed = false;

                for (int i = 0; i < DrillCount; i++)
                {
                    string currentValue = drillTextBoxes[i].Text.Trim();
                    string defaultValue = $"DRILL_{i + 1}";
                    bool isDrillOne = (i + 1) == 1;

                    // Determine if a set is needed
                    bool needsSet = false;

                    if (isDrillOne)
                    {
                        // For DRILL_1, set only if not already default
                        if (!currentValue.Equals(defaultValue, StringComparison.OrdinalIgnoreCase))
                        {
                            needsSet = true;
                        }
                    }
                    else
                    {
                        // For DRILL_2 to DRILL_12, set only if not default
                        if (!currentValue.Equals(defaultValue, StringComparison.OrdinalIgnoreCase))
                        {
                            needsSet = true;
                        }
                    }

                    if (needsSet)
                    {
                        // Perform the set operation for this drill
                        SetDrill(i);
                        anySetPerformed = true;
                    }
                }

                if (anySetPerformed)
                {
                    MessageBox.Show("All applicable drills have been set with your inputs.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("All drills are already at their original values. No set needed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        /// <summary>
        /// Handles the RESET ALL button click.
        /// </summary>
        private void ResetAllButton_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show("Are you sure you want to reset all drills?", "Confirm RESET ALL", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmResult == DialogResult.Yes)
            {
                bool anyResetPerformed = false;

                for (int i = 0; i < DrillCount; i++)
                {
                    string currentValue = drillTextBoxes[i].Text.Trim();
                    string defaultValue = $"DRILL_{i + 1}";
                    bool isDrillOne = (i + 1) == 1;
                    bool needsReset = false;

                    if (isDrillOne)
                    {
                        if (!currentValue.Equals(defaultValue, StringComparison.OrdinalIgnoreCase))
                        {
                            needsReset = true;
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrWhiteSpace(currentValue))
                        {
                            needsReset = true;
                        }
                    }

                    if (needsReset)
                    {
                        // Perform the reset operation for this drill
                        ResetDrill(i);
                        anyResetPerformed = true;
                    }
                }

                if (anyResetPerformed)
                {
                    MessageBox.Show("All applicable drills have been reset.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("All drills are already at their default values. No reset needed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        /// <summary>
        /// Handles the UPDATE FROM BLOCK ATTRIBUTE button click.
        /// </summary>
        private void UpdateFromAttributesButton_Click(object sender, EventArgs e)
        {
            UpdateFromBlockAttribute();
        }





        private void CreateAllButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < DrillCount; i++)
            {
                CreateButton_Click(i); // Simulate individual create click
            }
        }


        /// <summary>
        /// Updates the form fields based on block attributes in the selected blocks.
        /// </summary>
        private void UpdateFromBlockAttribute()
        {
            try
            {
                var acDoc = AcApplication.DocumentManager.MdiActiveDocument;
                var editor = acDoc.Editor;

                // Prompt for selection of blocks
                PromptSelectionOptions options = new PromptSelectionOptions();
                options.MessageForAdding = "Select blocks with DRILL attributes:";
                PromptSelectionResult result = editor.GetSelection(options);

                if (result.Status != PromptStatus.OK)
                {
                    MessageBox.Show("No blocks selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SelectionSet selectionSet = result.Value;
                var selectedObjects = selectionSet.GetObjectIds();

                {
                    foreach (var objectId in selectedObjects)
                    {
                        var blockRef = tr.GetObject(objectId, OpenMode.ForRead) as BlockReference;
                        if (blockRef == null || blockRef.IsErased)
                            continue; // Skip null or erased block references

                        // Iterate through the attributes of the block
                        foreach (ObjectId attrId in blockRef.AttributeCollection)
                        {
                            AttributeReference attrRef = tr.GetObject(attrId, OpenMode.ForRead) as AttributeReference;
                            if (attrRef != null && attrRef.Tag.StartsWith("DRILL_", StringComparison.OrdinalIgnoreCase))
                            {
                                // Extract the drill number from the tag, e.g., "DRILL_12" -> 12
                                string numberPart = attrRef.Tag.Substring(6); // Get the part after "DRILL_"
                                if (int.TryParse(numberPart, out int drillNumber))
                                {
                                    if (drillNumber >= 1 && drillNumber <= DrillCount)
                                    {
                                        int index = drillNumber - 1; // Convert to zero-based index
                                        string textValue = !string.IsNullOrEmpty(attrRef.TextString) ? attrRef.TextString : $"DRILL_{drillNumber}";
                                        drillTextBoxes[index].Text = textValue;
                                        drillLabels[index].Text = textValue; // Update label as well
                                    }
                                    else
                                    {
                                        // Handle unexpected drill numbers if necessary
                                        MessageBox.Show($"Unexpected drill number: {drillNumber}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                                else
                                {
                                    // Handle non-integer drill tags if necessary
                                    MessageBox.Show($"Invalid drill tag format: {attrRef.Tag}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                        }
                    }
                    tr.Commit();
                }

                MessageBox.Show("Updated fields from block attributes.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to update from block attributes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Inserts a block into the AutoCAD layout with the specified drill name and index.
        /// </summary>
        /// <param name="drillName">Name of the drill.</param>
        /// <param name="drillIndex">Index of the drill (1-based).</param>
        private void InsertBlockIntoLayout(string drillName, int drillIndex)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            {
                BlockTable blockTable = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;

                if (!blockTable.Has("DRILL_NAME_HEADING"))
                {
                    MessageBox.Show("Block 'DRILL_NAME_HEADING' does not exist in the drawing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ObjectId blockId = blockTable["DRILL_NAME_HEADING"];
                if (blockId != ObjectId.Null)
                {
                    BlockReference blockRef = new BlockReference(new Autodesk.AutoCAD.Geometry.Point3d(0, 0, 0), blockId);

                    BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    modelSpace.AppendEntity(blockRef);
                    tr.AddNewlyCreatedDBObject(blockRef, true);

                    foreach (ObjectId attId in blockRef.AttributeCollection)
                    {
                        AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                        if (attRef != null && attRef.Tag.Equals($"DRILL_{drillIndex}", StringComparison.OrdinalIgnoreCase))
                        {
                            attRef.TextString = drillName;
                        }
                    }
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Saves the current drill names to a JSON file.
        /// </summary>
        private void SaveToJson()
        {
            string jsonFilePath = GetJsonFilePath();

            // Ensure exactly DrillCount entries
            string[] drillNames = new string[DrillCount];
            for (int i = 0; i < DrillCount; i++)
            {
                drillNames[i] = drillTextBoxes[i].Text.Trim();
            }

            try
            {
                File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(drillNames, Formatting.Indented));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save JSON: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handles the SAVE DATA operation.
        /// </summary>
        private void SaveData()
        {
            SaveToJson();
        }

        /// <summary>
        /// Loads drill names from a JSON file.
        /// </summary>
        private void LoadFromJson()
        {
            try
            {
                string jsonFilePath = GetJsonFilePath();

                if (!File.Exists(jsonFilePath))
                {
                    // Initialize with default drill names
                    InitializeDefaultDrills();

                    // Save default drills to JSON
                    SaveToJson();

                    MessageBox.Show($"JSON file not found. Initialized with default drill names and created '{Path.GetFileName(jsonFilePath)}'.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string jsonData = File.ReadAllText(jsonFilePath);
                var drillNames = JsonConvert.DeserializeObject<string[]>(jsonData);

                if (drillNames == null)
                {
                    MessageBox.Show("JSON data is null. Initializing with default drill names.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    InitializeDefaultDrills();
                    SaveToJson();
                    return;
                }

                // Validate and load drill names
                for (int i = 0; i < DrillCount; i++)
                {
                    if (i < drillNames.Length && !string.IsNullOrWhiteSpace(drillNames[i]))
                    {
                        drillTextBoxes[i].Text = drillNames[i].Trim();
                        drillLabels[i].Text = drillNames[i].Trim();  // Load labels as well from the JSON
                    }
                    else
                    {
                        // If JSON has fewer entries or entry is empty, initialize remaining drills with default names
                        string defaultName = $"DRILL_{i + 1}";
                        drillTextBoxes[i].Text = defaultName;
                        drillLabels[i].Text = defaultName;
                    }
                }

                // Optionally, verify if all drills have been loaded correctly
                // You can add a confirmation message here if needed

                // Save back the potentially updated drill names
                SaveToJson();
            }
            catch (JsonException jex)
            {
                MessageBox.Show($"Failed to parse JSON: {jex.Message}", "JSON Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                InitializeDefaultDrills();
                SaveToJson();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load JSON: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Initializes default drill names in the form.
        /// </summary>
        private void InitializeDefaultDrills()
        {
            for (int i = 0; i < DrillCount; i++)
            {
                string defaultName = $"DRILL_{i + 1}";
                drillTextBoxes[i].Text = defaultName;
                drillLabels[i].Text = defaultName;
            }
        }
    }
}

namespace Drill_Namer.Models
{
    public class BlockAttributeData
    {
        public BlockReference BlockReference { get; set; }
        public string DrillName { get; set; }
        public double YCoordinate { get; set; }
    }
}
namespace Drill_Namer.Models
{
    public class BlockData
    {
        public int WellId { get; set; }
        public string DrillName { get; set; }
    }
}

namespace Drill_Namer.Models
{
    /// <summary>
    /// Represents the coordinates data for a drill.
    /// </summary>
    public class CoordinatesData
    {
        [JsonProperty("wellbore")]
        public List<string> Wellbore { get; set; }

        [JsonProperty("leg_#")]
        public List<int> LegNumbers { get; set; }

        [JsonProperty("order")]
        public List<int> Order { get; set; }

        [JsonProperty("latitude")]
        public List<string> Latitudes { get; set; }    // Changed from List<double> to List<string>

        [JsonProperty("longitude")]
        public List<string> Longitudes { get; set; }   // Changed from List<double> to List<string>
    }
}

namespace Drill_Namer.Models
{
    public class DrillData
    {
        [JsonProperty("gw_id")]
        public int WellId { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("file_#")]
        public string FileNumber { get; set; }

        [JsonProperty("final_survey_date")]
        public string FinalSurveyDate { get; set; }

        [JsonProperty("revision_#")]
        public int RevisionNumber { get; set; }

        [JsonProperty("as_drilled")]
        public bool AsDrilled { get; set; }

        [JsonProperty("lateral_len")]
        public double LateralLength { get; set; }

        [JsonProperty("coordinates")]
        public CoordinatesData Coordinates { get; set; }
    }
}

namespace Drill_Namer.Models
{
    public class WellsData
    {
        [JsonProperty("wells")]
        public List<DrillData> Wells { get; set; } = new List<DrillData>();
    }
}
