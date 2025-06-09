// NOTE: Ensure the project is targeting x64 in Visual Studio to match AutoCAD DLLs.
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Drawing;
using OfficeOpenXml;
using Newtonsoft.Json;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Windows.Data;
using NLog;
using NLog.Config;
using NLog.Targets;
using Drill_Namer.Models;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using AColor = Autodesk.AutoCAD.Colors.Color;
using DrawingColor = System.Drawing.Color;
using FormsFlowDirection = System.Windows.Forms.FlowDirection;

namespace Drill_Namer
{
    public static class AutoCADHelper
    {
        /// <summary>
        /// Start a transaction in the current AutoCAD document.
        /// </summary>
        public static Transaction StartTransaction()
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
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
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
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
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            ObjectId brId = ObjectId.Null;

            // Make sure we lock the doc for changes
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
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
            using (Transaction tr = db.TransactionManager.StartTransaction())
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
            using (Database tempDb = new Database(false, true))
            {
                try
                {
                    // read the .dwg containing the block
                    tempDb.ReadDwgFile(sourceDwgPath, FileShare.Read, true, "");
                    using (Transaction tr = tempDb.TransactionManager.StartTransaction())
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
                    using (Transaction destTr = destDb.TransactionManager.StartTransaction())
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
            using (Transaction tr = db.TransactionManager.StartTransaction())
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
            catch (System.Exception ex)
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
                Document acDoc = AcApplication.DocumentManager.MdiActiveDocument;
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
    public partial class FindReplaceForm : Form
    {
        // Static constructor registers the AssemblyResolve event handler
        static FindReplaceForm()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        // This event handler attempts to load System.Text.Encoding.CodePages.dll
        // from the same folder as your plugin.
        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            AssemblyName requestedName = new AssemblyName(args.Name);
            if (requestedName.Name.Equals("System.Text.Encoding.CodePages", StringComparison.OrdinalIgnoreCase))
            {
                // Get the directory of the executing assembly (your plugin)
                string pluginFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string dllPath = Path.Combine(pluginFolder, "System.Text.Encoding.CodePages.dll");
                if (File.Exists(dllPath))
                {
                    return Assembly.LoadFrom(dllPath);
                }
            }
            return null;
        }

        #region UI Initialization

        // UI component declarations
        private int DrillCount;
        private TextBox[] drillTextBoxes;
        private Label[] drillLabels;
        private Button[] headingButtons;
        private Button[] setButtons;
        private Button[] resetButtons;
        private ComboBox headingComboBox;
        private Button updateFromAttributesButton;
        private ComboBox swapComboBox1;
        private ComboBox swapComboBox2;
        private Button swapButton;
        private Button headingAllButton;
        private Button setAllButton;
        private Button resetAllButton;
        private Button checkButton;
        private GroupBox drillGroupBox;
        private ComboBox drillComboBox;
        private Button generateJsonButton;

        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ClientSize = new System.Drawing.Size(600, 650);
            this.Name = "FindReplaceForm";
            this.Text = "Find and Replace Form";
            this.ResumeLayout(false);
        }

        public FindReplaceForm()
        {
            InitializeComponent();
            // Register the CodePagesEncodingProvider (needed by EPPlus and similar libraries)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            this.Text = "DRILL PROPERTIES";
            DrillCount = 12; // For example

            InitializeDynamicControls();
            LoadFromJson();
            Logger.LogInfo("FindReplaceForm initialized.");
            Logger.LogInfo("Form constructor executed successfully.");
        }

        private void InitializeDynamicControls()
        {
            // Basic form settings
            // Let the user resize the window but show scrollbars when needed
            this.AutoScroll = true;
            this.AutoSize = false;
            this.BackColor = System.Drawing.Color.Black;
            this.ForeColor = System.Drawing.Color.White;
            this.Text = "DRILL PROPERTIES";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(900, 600);

            // Arrays for dynamic controls
            drillTextBoxes = new TextBox[DrillCount];
            drillLabels = new Label[DrillCount];
            headingButtons = new Button[DrillCount];
            setButtons = new Button[DrillCount];
            resetButtons = new Button[DrillCount];

            // Example fonts
            var labelFont = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold);
            var textBoxFont = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular);
            var buttonFont = new System.Drawing.Font("Segoe UI", 8, System.Drawing.FontStyle.Bold);

            // Main scrolling panel
            Panel mainPanel = new Panel
            {
                Name = "mainPanel",
                AutoScroll = true,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                Dock = DockStyle.Fill
            };
            this.Controls.Add(mainPanel);

            // GroupBox for drill controls
            drillGroupBox = new GroupBox
            {
                Text = "Drill Controls",
                Location = new System.Drawing.Point(10, 10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                Padding = new Padding(10)
            };
            mainPanel.Controls.Add(drillGroupBox);

            // TableLayoutPanel for DRILL Label/Name/Heading/Set/Reset columns
            TableLayoutPanel tableLayout = new TableLayoutPanel
            {
                Location = new System.Drawing.Point(10, 20),
                ColumnCount = 5,
                RowCount = DrillCount + 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                Dock = DockStyle.Top,
                Padding = new Padding(5)
            };

            // Give labels and names more room
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));  // Drill Label
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));  // Drill Name
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 8F));   // HEADING (wider so text isn't cut off)
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 6F));   // SET
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 6F));   // RESET

            // Header row
            tableLayout.Controls.Add(new Label
            {
                Text = "Drill Label",
                Font = labelFont,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            }, 0, 0);
            tableLayout.Controls.Add(new Label
            {
                Text = "Drill Name",
                Font = labelFont,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            }, 1, 0);
            tableLayout.Controls.Add(new Label
            {
                Text = "HEADING",
                Font = labelFont,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            }, 2, 0);
            tableLayout.Controls.Add(new Label
            {
                Text = "SET",
                Font = labelFont,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            }, 3, 0);
            tableLayout.Controls.Add(new Label
            {
                Text = "RESET",
                Font = labelFont,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            }, 4, 0);

            // Populate drill rows
            for (int i = 0; i < DrillCount; i++)
            {
                // DRILL Label
                drillLabels[i] = new Label
                {
                    Text = $"DRILL_{i + 1}",
                    Font = labelFont,
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                    Dock = DockStyle.Fill,
                    BackColor = System.Drawing.Color.Black,
                    ForeColor = System.Drawing.Color.White
                };
                tableLayout.Controls.Add(drillLabels[i], 0, i + 1);

                // DRILL TextBox
                drillTextBoxes[i] = new TextBox
                {
                    Font = textBoxFont,
                    Dock = DockStyle.Fill,
                    BackColor = System.Drawing.Color.Black,
                    ForeColor = System.Drawing.Color.White
                };
                tableLayout.Controls.Add(drillTextBoxes[i], 1, i + 1);

                // Let the textbox grow/shrink with the form
                drillTextBoxes[i].Anchor = AnchorStyles.Left | AnchorStyles.Right;
                drillTextBoxes[i].MaximumSize = new Size(0, 20);   // (0 = no horizontal limit)

                // HEADING button
                headingButtons[i] = new Button
                {
                    Text = "HEADING",
                    Font = buttonFont,
                    BackColor = System.Drawing.Color.DarkGreen,
                    ForeColor = System.Drawing.Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Dock = DockStyle.Fill
                };
                int rowIndex = i;
                headingButtons[i].Click += (sender, e) => HeadingButton_Click(rowIndex);
                tableLayout.Controls.Add(headingButtons[i], 2, i + 1);

                // SET button
                setButtons[i] = new Button
                {
                    Text = "SET",
                    Font = buttonFont,
                    BackColor = System.Drawing.Color.DarkCyan,
                    ForeColor = System.Drawing.Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Dock = DockStyle.Fill
                };
                setButtons[i].Click += (sender, e) => SetDrill(rowIndex);
                tableLayout.Controls.Add(setButtons[i], 3, i + 1);

                // RESET button
                resetButtons[i] = new Button
                {
                    Text = "RESET",
                    Font = buttonFont,
                    BackColor = System.Drawing.Color.DarkRed,
                    ForeColor = System.Drawing.Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Dock = DockStyle.Fill
                };
                resetButtons[i].Click += (sender, e) => ResetDrill(rowIndex);
                tableLayout.Controls.Add(resetButtons[i], 4, i + 1);
            }

            drillGroupBox.Controls.Add(tableLayout);
            int nextY = tableLayout.Bottom + 20;

            // Drill selection ComboBox
            drillComboBox = new ComboBox
            {
                Name = "drillComboBox",
                Location = new System.Drawing.Point(10, nextY),
                Size = new System.Drawing.Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            };
            drillComboBox.Items.Add("ALL");
            for (int i = 0; i < DrillCount; i++)
            {
                drillComboBox.Items.Add($"DRILL_{i + 1}");
            }
            drillComboBox.SelectedIndex = 0;
            drillGroupBox.Controls.Add(drillComboBox);

            // Generate JSON button
            generateJsonButton = new Button
            {
                Text = "Generate JSON",
                Location = new System.Drawing.Point(drillComboBox.Right + 10, nextY),
                Size = new System.Drawing.Size(120, 25),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Gold,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            generateJsonButton.Click += GenerateJsonButton_Click;
            drillGroupBox.Controls.Add(generateJsonButton);

            nextY = drillComboBox.Bottom + 20;

            // Heading type ComboBox
            headingComboBox = new ComboBox()
            {
                Location = new System.Drawing.Point(10, nextY),
                Size = new System.Drawing.Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            };
            headingComboBox.Items.AddRange(new string[] { "VEREN", "OTHER" });
            headingComboBox.SelectedIndex = 0;
            headingComboBox.SelectedIndexChanged += (sender, e) => SaveToJson();

            Label clientLabel = new Label()
            {
                Text = "CLIENT",
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                AutoSize = true,
                Location = new System.Drawing.Point(10, nextY)
            };
            int clientLabelWidth = TextRenderer.MeasureText(clientLabel.Text, clientLabel.Font).Width;
            headingComboBox.Location = new System.Drawing.Point(clientLabel.Location.X + clientLabelWidth + 10, nextY);

            drillGroupBox.Controls.Add(clientLabel);
            drillGroupBox.Controls.Add(headingComboBox);

            nextY = headingComboBox.Bottom + 20;

            // UPDATE FROM BLOCK ATTRIBUTE button
            updateFromAttributesButton = new Button()
            {
                Text = "UPDATE FROM BLOCK ATTRIBUTE",
                Location = new System.Drawing.Point(10, nextY),
                Size = new System.Drawing.Size(220, 30),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Gray,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            updateFromAttributesButton.Click += UpdateFromAttributesButton_Click;
            drillGroupBox.Controls.Add(updateFromAttributesButton);

            nextY = updateFromAttributesButton.Bottom + 20;

            // SWAP ComboBoxes
            swapComboBox1 = new ComboBox()
            {
                Location = new System.Drawing.Point(10, nextY),
                Size = new System.Drawing.Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            };
            swapComboBox2 = new ComboBox()
            {
                Location = new System.Drawing.Point(swapComboBox1.Right + 10, nextY),
                Size = new System.Drawing.Size(100, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White
            };
            for (int i = 0; i < DrillCount; i++)
            {
                swapComboBox1.Items.Add($"DRILL_{i + 1}");
                swapComboBox2.Items.Add($"DRILL_{i + 1}");
            }
            swapComboBox1.SelectedIndex = 0;
            swapComboBox2.SelectedIndex = 1;
            drillGroupBox.Controls.Add(swapComboBox1);
            drillGroupBox.Controls.Add(swapComboBox2);

            // SWAP button
            swapButton = new Button()
            {
                Text = "SWAP",
                Location = new System.Drawing.Point(swapComboBox2.Right + 10, nextY),
                Size = new System.Drawing.Size(75, 25),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Orange,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            swapButton.Click += SwapButton_Click;
            drillGroupBox.Controls.Add(swapButton);

            // Optional logo
            int logoHeight = 80;
            int logoWidth = 200;
            int swapBottom = swapButton.Top + swapButton.Height;
            int logoY = swapBottom - logoHeight;
            int logoX = swapButton.Right + 10;
            string logoPath = @"C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties\CompassLogo.png";
            PictureBox logoPictureBox = new PictureBox()
            {
                Name = "logoPictureBox",
                Size = new System.Drawing.Size(logoWidth, logoHeight),
                Location = new System.Drawing.Point(logoX, logoY),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                BackColor = System.Drawing.Color.Black
            };
            if (System.IO.File.Exists(logoPath))
            {
                logoPictureBox.Image = System.Drawing.Image.FromFile(logoPath);
            }
            else
            {
                Logger.LogWarning($"Logo not found at {logoPath}");
            }
            drillGroupBox.Controls.Add(logoPictureBox);
            logoPictureBox.BringToFront();

            // Bottom panel with vertical stacking (TopDown)
            FlowLayoutPanel bottomPanel = new FlowLayoutPanel()
            {
                FlowDirection = System.Windows.Forms.FlowDirection.TopDown,  // <-- Fully qualify
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Location = new System.Drawing.Point(10, Math.Max(swapButton.Bottom, logoPictureBox.Bottom) + 15),
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                Margin = new Padding(0, 10, 0, 0)
            };
            drillGroupBox.Controls.Add(bottomPanel);

            // ROW 1: SET ALL, RESET ALL
            FlowLayoutPanel row1 = new FlowLayoutPanel()
            {
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight, // <-- Fully qualify
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            setAllButton = new Button()
            {
                Text = "SET ALL",
                Size = new System.Drawing.Size(100, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.DarkCyan,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            setAllButton.Click += SetAllButton_Click;
            row1.Controls.Add(setAllButton);

            resetAllButton = new Button()
            {
                Text = "RESET ALL",
                Size = new System.Drawing.Size(100, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.DarkRed,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            resetAllButton.Click += ResetAllButton_Click;
            row1.Controls.Add(resetAllButton);

            bottomPanel.Controls.Add(row1);

            // ROW 2: CHECK, WELL CORNERS, HEADING ALL
            FlowLayoutPanel row2 = new FlowLayoutPanel()
            {
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight, // <-- Fully qualify
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            checkButton = new Button()
            {
                Text = "CHECK",
                Size = new System.Drawing.Size(100, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.Yellow,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            checkButton.Click += CheckButton_Click;
            row2.Controls.Add(checkButton);

            Button wellCornersButton = new Button()
            {
                Text = "WELL CORNERS",
                Size = new System.Drawing.Size(120, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.MediumPurple,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            wellCornersButton.Click += WellCornersButton_Click;
            row2.Controls.Add(wellCornersButton);

            headingAllButton = new Button()
            {
                Text = "HEADING ALL",
                Size = new System.Drawing.Size(100, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.DarkGreen,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            headingAllButton.Click += HeadingAllButton_Click;
            row2.Controls.Add(headingAllButton);

            bottomPanel.Controls.Add(row2);

            // ROW 3: CREATE TABLE, CREATE XLS, COMPLETE CORDS
            FlowLayoutPanel row3 = new FlowLayoutPanel()
            {
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight, // <-- Fully qualify
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            Button createTableButton = new Button()
            {
                Text = "CREATE TABLE",
                Size = new System.Drawing.Size(120, 30),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.LightPink,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            createTableButton.Click += CreateTableButton_Click;
            row3.Controls.Add(createTableButton);

            Button createXlsButton = new Button()
            {
                Text = "CREATE XLS",
                Size = new System.Drawing.Size(100, 30),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.LightPink,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            createXlsButton.Click += CreateXlsButton_Click;
            row3.Controls.Add(createXlsButton);

            Button completeCordsButton = new Button()
            {
                Text = "COMPLETE CORDS",
                Size = new System.Drawing.Size(140, 30),
                Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.Orange,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            completeCordsButton.Click += CompleteCordsButton_Click;
            row3.Controls.Add(completeCordsButton);

            bottomPanel.Controls.Add(row3);

            // ROW 4: GET UTMS, ADD DRILL PTS
            FlowLayoutPanel row4 = new FlowLayoutPanel()
            {
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight, // <-- Fully qualify
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            Button getUtmsButton = new Button()
            {
                Text = "GET UTMS",
                Size = new System.Drawing.Size(100, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.LightBlue,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            getUtmsButton.Click += GetUtmsButton_Click;
            row4.Controls.Add(getUtmsButton);

            Button addDrillPtsButton = new Button()
            {
                Text = "ADD DRILL PTS",
                Size = new System.Drawing.Size(120, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.LightBlue,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            addDrillPtsButton.Click += AddDrillPtsButton_Click;
            row4.Controls.Add(addDrillPtsButton);

            bottomPanel.Controls.Add(row4);

            // ROW 5: NEW: UPDATE OFFSETS
            FlowLayoutPanel row5 = new FlowLayoutPanel()
            {
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            Button updateOffsetsButton = new Button()
            {
                Text = "UPDATE OFFSETS",
                Size = new System.Drawing.Size(140, 30),
                Font = buttonFont,
                BackColor = System.Drawing.Color.LightGreen,
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = FlatStyle.Flat
            };
            updateOffsetsButton.Click += UpdateOffsetsButton_Click;
            row5.Controls.Add(updateOffsetsButton);
            bottomPanel.Controls.Add(row5);

            // Status strip
            StatusStrip statusStrip = new StatusStrip()
            {
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                Dock = DockStyle.Bottom
            };
            ToolStripStatusLabel statusLabel = new ToolStripStatusLabel()
            {
                Text = "Ready",
                Spring = true,
                ForeColor = System.Drawing.Color.White
            };
            statusStrip.Items.Add(statusLabel);
            drillGroupBox.Controls.Add(statusStrip);

            // Give the form a generous starting size
            this.ClientSize = new System.Drawing.Size(1600, 800);
            this.DoubleBuffered = true;
        }

        #endregion



        private DrillData GenerateJsonForDrill(int drillIndex, Table coordinateTable = null, bool collectDataOnly = false)
        {
            string drillName = string.Empty;
            try
            {
                // Step 1: Extract wellId and drillName from block
                BlockData blockData = GetBlockData();
                if (blockData == null)
                {
                    // User cancelled or an error occurred
                    return null;
                }
                int wellId = blockData.WellId;
                drillName = blockData.DrillName;

                // Step 2: Get fileNumber
                string fileNumber = GetDrawingName();

                // Step 3: Get final_survey_date and revisionNumber
                GetRevisionInfo(out int revisionNumber, out string finalSurveyDate);

                // Step 4: as_drilled (always false)
                bool asDrilled = false;

                // Step 5: Calculate lateral_len
                double lateralLength = GetLateralLengthFromUser();
                if (double.IsNaN(lateralLength))
                {
                    // User cancelled or an error occurred
                    return null;
                }

                // Step 6: Build coordinates
                CoordinatesData coordinates = GetCoordinatesForDrill(drillIndex, coordinateTable);
                if (coordinates == null)
                {
                    // Handle error
                    Logger.LogError($"Failed to get coordinates for drill index '{drillIndex + 1}'");
                    return null;
                }

                // Step 7: Compile data into DrillData object
                DrillData drillData = new DrillData
                {
                    WellId = wellId,
                    Name = drillName,
                    FileNumber = fileNumber,
                    FinalSurveyDate = finalSurveyDate,
                    RevisionNumber = revisionNumber,
                    AsDrilled = asDrilled,
                    LateralLength = lateralLength,
                    Coordinates = coordinates
                };

                if (collectDataOnly)
                {
                    return drillData;
                }

                // Step 8: Save JSON to file
                SaveJsonToFile(drillData, drillName);

                MessageBox.Show($"JSON file generated successfully for {drillName}.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo($"JSON file generated for drill '{drillName}'.");

                return null;
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error generating JSON for drill '{drillName}': {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while generating JSON for drill '{drillName}': {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        
        private CoordinatesData GetCoordinatesForDrill(int drillIndex, Table coordinateTable)
        {
            try
            {
                if (coordinateTable == null)
                {
                    // If the table is not provided, prompt the user to select it
                    coordinateTable = SelectCoordinateTable();
                    if (coordinateTable == null)
                    {
                        Logger.LogInfo("User cancelled table selection.");
                        return null;
                    }
                }

                // Extract data for the specified drill index
                CoordinatesData coordData = ExtractCoordinatesFromTable(coordinateTable, drillIndex);

                return coordData;
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetCoordinatesForDrill: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while retrieving coordinates: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        private CoordinatesData ExtractCoordinatesFromTable(Table table, int drillIndex)
        {
            List<string> wellbore = new List<string>();
            List<int> legNumbers = new List<int>();
            List<int> order = new List<int>();
            List<string> latitudes = new List<string>();    // Changed to List<string>
            List<string> longitudes = new List<string>();   // Changed to List<string>

            int currentDrillCount = -1; // Start at -1 because we increment before checking
            bool collectingData = false;

            Logger.LogInfo($"Starting to extract coordinates for drill index '{drillIndex}'");

            for (int row = 0; row < table.Rows.Count; row++)
            {
                // Read the value in column A (column index 0)
                string columnAValue = table.Cells[row, 0].Value?.ToString().Trim() ?? "";

                // Remove formatting codes if necessary
                columnAValue = RemoveFormattingCodes(columnAValue);

                // Normalize columnAValue
                string normalizedColumnAValue = Regex.Replace(columnAValue.Trim().ToUpper(), @"\s+", "");

                Logger.LogInfo($"Row {row}: Column A value '{columnAValue}', normalized '{normalizedColumnAValue}'");

                // Check for the "SURFACE" marker to identify the start of a drill section
                if (normalizedColumnAValue.Contains("SURFACE"))
                {
                    currentDrillCount++;

                    if (currentDrillCount == drillIndex)
                    {
                        collectingData = true;
                        Logger.LogInfo($"Found drill index '{drillIndex}' at row {row}. Starting data collection.");
                        // Collect data for the current row (the "SURFACE" point)
                        CollectCoordinateData(table, row, wellbore, legNumbers, order, latitudes, longitudes);
                        continue;
                    }
                    else if (collectingData)
                    {
                        // We've reached the next "SURFACE" marker, so stop collecting
                        collectingData = false;
                        Logger.LogInfo($"End of data collection for drill index '{drillIndex}' at row {row}.");
                        break;
                    }
                }

                if (collectingData)
                {
                    // Collect data for the current row
                    CollectCoordinateData(table, row, wellbore, legNumbers, order, latitudes, longitudes);
                }
            }

            if (wellbore.Count == 0)
            {
                MessageBox.Show($"No coordinate data found for drill index '{drillIndex + 1}'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError($"No coordinate data found for drill index '{drillIndex + 1}'");
                return null;
            }

            Logger.LogInfo($"Successfully extracted coordinates for drill index '{drillIndex}'");

            return new CoordinatesData
            {
                Wellbore = wellbore,
                LegNumbers = legNumbers,
                Order = order,
                Latitudes = latitudes,
                Longitudes = longitudes
            };
        }
        private double GetLateralLengthFromUser()
        {
            try
            {
                // Use the custom input dialog instead of Interaction.InputBox
                string input = ShowInputDialog("Enter lateral length:", "Lateral Length", "0");

                if (double.TryParse(input, out double lateralLength) && lateralLength >= 0)
                {
                    Logger.LogInfo($"User entered lateral length: {lateralLength}");
                    return lateralLength;
                }
                else
                {
                    MessageBox.Show("Please enter a valid non-negative number for lateral length.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Logger.LogWarning($"Invalid lateral length input: '{input}'");
                    return double.NaN;
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetLateralLengthFromUser: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while getting lateral length: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return double.NaN;
            }
        }
        private void SaveJsonToFile(DrillData jsonData, string drillName)
        {
            try
            {
                // Serialize to JSON without indentation (single line)
                string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.None);

                // Prompt user to select save location
                System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog
                {
                    FileName = $"{drillName}_Data.json",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    DefaultExt = "json",
                    AddExtension = true
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Write the JSON string to the file
                    File.WriteAllText(saveFileDialog.FileName, jsonString);
                    Logger.LogInfo($"JSON file saved to {saveFileDialog.FileName}");
                }
                else
                {
                    Logger.LogInfo("User cancelled JSON save dialog.");
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error saving JSON file: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while saving the JSON file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string RemoveFormattingCodes(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            try
            {
                // Remove RTF formatting codes using regex
                // Matches any content starting with '{' and ending with ';', including the braces and semicolon
                text = Regex.Replace(text, @"\{.*?;", "");

                // Remove any closing braces '}' that might remain
                text = text.Replace("}", "");

                return text;
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in RemoveFormattingCodes: {ex.Message}\n{ex.StackTrace}");
                return text; // Return original text if an error occurs
            }
        }
        private void GetRevisionInfo(out int revisionNumber, out string finalSurveyDate)
        {
            revisionNumber = 0;

            try
            {
                string drawingName = GetDrawingName();
                revisionNumber = ParseRevisionNumber(drawingName);
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetRevisionInfo: {ex.Message}\n{ex.StackTrace}");
            }

            // Set finalSurveyDate to the current date in MM/dd/yy format with slashes
            finalSurveyDate = DateTime.Now.ToString("MM'/'dd'/'yy");

            // Log the retrieved values
            Logger.LogInfo($"Revision Number: {revisionNumber}");
            Logger.LogInfo($"Final Survey Date: {finalSurveyDate}");
        }
        private BlockData GetBlockData()
        {
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                // Prompt the user to select a block reference
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect the block with WELLID and DRILLNAME attributes:");
                peo.SetRejectMessage("\nOnly block references are allowed.");
                peo.AddAllowedClass(typeof(BlockReference), true);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nSelection cancelled.");
                    Logger.LogInfo("User cancelled block selection.");
                    return null;
                }

                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    BlockReference blkRef = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;

                    if (blkRef != null)
                    {
                        int wellId = -1;
                        string drillName = null;

                        foreach (ObjectId attId in blkRef.AttributeCollection)
                        {
                            AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;

                            if (attRef != null)
                            {
                                if (attRef.Tag.Equals("WELLID", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Use regex to extract the first integer from the WELLID value
                                    Match match = Regex.Match(attRef.TextString, @"\d+");
                                    if (match.Success && int.TryParse(match.Value, out int id))
                                    {
                                        wellId = id;
                                        Logger.LogInfo($"Extracted WELLID: {wellId}");
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Invalid WELLID value: {attRef.TextString}. It should contain an integer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        Logger.LogError($"Invalid WELLID value: {attRef.TextString}");
                                        return null;
                                    }
                                }
                                else if (attRef.Tag.Equals("DRILLNAME", StringComparison.OrdinalIgnoreCase))
                                {
                                    drillName = attRef.TextString.Trim();
                                    Logger.LogInfo($"Extracted DRILLNAME: {drillName}");
                                }
                            }
                        }

                        tr.Commit();

                        if (wellId == -1 || string.IsNullOrEmpty(drillName))
                        {
                            MessageBox.Show("Could not retrieve WELLID or DRILLNAME from the selected block.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Logger.LogError("Missing WELLID or DRILLNAME in selected block.");
                            return null;
                        }

                        return new BlockData
                        {
                            WellId = wellId,
                            DrillName = drillName
                        };
                    }
                    else
                    {
                        MessageBox.Show("Selected entity is not a valid block reference.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Logger.LogError("Selected entity is not a BlockReference.");
                        return null;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetBlockData: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while retrieving block data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        private void CollectCoordinateData(Table table, int row, List<string> wellbore, List<int> legNumbers, List<int> order, List<string> latitudes, List<string> longitudes)
        {
            // Read point type from column A (index 0)
            string pointType = table.Cells[row, 0].Value?.ToString().Trim() ?? "";
            pointType = RemoveFormattingCodes(pointType);

            // Read latitude and longitude from columns G and H (indices 6 and 7)
            string latitudeStr = table.Cells[row, 6].Value?.ToString().Trim() ?? ""; // Column G
            string longitudeStr = table.Cells[row, 7].Value?.ToString().Trim() ?? ""; // Column H

            // Remove formatting codes if necessary
            latitudeStr = RemoveFormattingCodes(latitudeStr);
            longitudeStr = RemoveFormattingCodes(longitudeStr);

            Logger.LogInfo($"Row {row}: PointType '{pointType}', Latitude '{latitudeStr}', Longitude '{longitudeStr}'");

            // Parse latitude and longitude
            if (double.TryParse(latitudeStr, out double latitude) && double.TryParse(longitudeStr, out double longitude))
            {
                // Determine wellbore point
                string wellborePoint = "";
                string normalizedPointType = Regex.Replace(pointType.ToUpper(), @"\s+", "");

                if (normalizedPointType.Contains("SURFACE"))
                    wellborePoint = "origin";
                else if (normalizedPointType.Contains("HEEL"))
                    wellborePoint = "heel_point";
                else if (normalizedPointType.Contains("BOTTOMHOLE"))
                    wellborePoint = "bottom_hole";
                else
                    wellborePoint = "sidetrack_point";

                wellbore.Add(wellborePoint);
                legNumbers.Add(1); // Assuming leg number is always 1
                order.Add(order.Count + 1);
                latitudes.Add(latitude.ToString("F6"));    // Format to 6 decimal places
                longitudes.Add(longitude.ToString("F6"));  // Format to 6 decimal places

                Logger.LogInfo($"Added data: Wellbore '{wellborePoint}', Latitude {latitude.ToString("F6")}, Longitude {longitude.ToString("F6")}");
            }
            else
            {
                Logger.LogWarning($"Failed to parse latitude or longitude at row {row}. Latitude string: '{latitudeStr}', Longitude string: '{longitudeStr}'");
            }
        }
        private List<BlockAttributeData> SelectAndExtractBlockData()
        {
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                // Prompt the user to select blocks
                PromptSelectionOptions pso = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect blocks with 'DRILLNAME' attribute:"
                };
                pso.AllowDuplicates = false;

                SelectionFilter filter = new SelectionFilter(new TypedValue[]
                {
            new TypedValue((int)DxfCode.Start, "INSERT") // Blocks
                });

                PromptSelectionResult psr = ed.GetSelection(pso, filter);

                if (psr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nBlock selection cancelled.");
                    Logger.LogInfo("User cancelled block selection.");
                    return null;
                }

                SelectionSet ss = psr.Value;
                List<BlockAttributeData> blockDataList = new List<BlockAttributeData>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in ss)
                    {
                        if (selObj != null)
                        {
                            BlockReference blockRef = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as BlockReference;
                            if (blockRef != null)
                            {
                                // Get the DRILLNAME attribute value
                                string drillNameAttribute = GetAttributeValue(blockRef, "DRILLNAME", tr);

                                if (!string.IsNullOrEmpty(drillNameAttribute))
                                {
                                    // Get the Y-coordinate
                                    double yCoord = blockRef.Position.Y;

                                    // Add to the list
                                    blockDataList.Add(new BlockAttributeData
                                    {
                                        BlockReference = blockRef,
                                        DrillName = drillNameAttribute,
                                        YCoordinate = yCoord
                                    });
                                }
                            }
                        }
                    }

                    tr.Commit();
                }

                Logger.LogInfo($"Extracted {blockDataList.Count} blocks with 'DRILLNAME' attribute.");
                return blockDataList;
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in SelectAndExtractBlockData: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while selecting and extracting block data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        private string GetDrawingName()
        {
            try
            {
                Document acDoc = AcApplication.DocumentManager.MdiActiveDocument;
                if (acDoc == null)
                {
                    Logger.LogError("No active AutoCAD document found.");
                    return string.Empty;
                }

                // Extract the file name without the path
                string fullPath = acDoc.Name;
                string fileName = System.IO.Path.GetFileNameWithoutExtension(fullPath);
                Logger.LogInfo($"Retrieved Drawing Name: {fileName}");
                return fileName;
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetDrawingName: {ex.Message}\n{ex.StackTrace}");
                return string.Empty;
            }
        }
        private int ParseRevisionNumber(string drawingName)
        {
            int revisionNumber = 0; // Default value

            try
            {
                if (string.IsNullOrEmpty(drawingName))
                {
                    Logger.LogWarning("Drawing name is empty. Defaulting revision number to 0.");
                    return revisionNumber;
                }

                // Find all matches of "-R" followed by digits
                string pattern = @"-R(\d+)";
                var matches = Regex.Matches(drawingName, pattern, RegexOptions.IgnoreCase);

                if (matches.Count > 0)
                {
                    // Get the last match to capture the latest revision
                    var lastMatch = matches[matches.Count - 1];
                    if (lastMatch.Groups.Count > 1)
                    {
                        string revStr = lastMatch.Groups[1].Value;
                        if (int.TryParse(revStr, out int revNum))
                        {
                            revisionNumber = revNum;
                            Logger.LogInfo($"Parsed Revision Number: {revisionNumber} from Drawing Name: {drawingName}");
                        }
                        else
                        {
                            Logger.LogError($"Failed to parse revision number from matched string: {revStr}. Defaulting to 0.");
                        }
                    }
                }
                else
                {
                    Logger.LogWarning($"No revision pattern found in Drawing Name: {drawingName}. Defaulting revision number to 0.");
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in ParseRevisionNumber: {ex.Message}\n{ex.StackTrace}");
            }

            return revisionNumber;
        }
        private void GenerateJsonButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (drillComboBox == null)
                {
                    MessageBox.Show("Drill ComboBox not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Logger.LogError("drillComboBox control is null.");
                    return;
                }

                if (drillComboBox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a drill or 'ALL' from the dropdown.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Logger.LogWarning("No selection made in drillComboBox.");
                    return;
                }

                if (drillComboBox.SelectedIndex == 0)
                {
                    HandleAllSelection();
                }
                else
                {
                    int drillIndex = drillComboBox.SelectedIndex - 1;
                    HandleSingleSelection(drillIndex);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError($"Exception in GenerateJsonButton_Click: {ex.Message}\n{ex.StackTrace}");
            }
        }
        private void HandleAllSelection()
        {
            try
            {
                // Step 1: Select the table once
                Table coordinateTable = SelectCoordinateTable();
                if (coordinateTable == null)
                {
                    MessageBox.Show("Table selection cancelled or invalid.", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                List<DrillData> allDrillData = new List<DrillData>();

                for (int i = 0; i < DrillCount; i++)
                {
                    string drillName = drillTextBoxes[i].Text.Trim();

                    // Generate the default drill name
                    string defaultDrillName = $"DRILL_{i + 1}";

                    // Skip drills that are at their default value or empty
                    if (string.IsNullOrEmpty(drillName) || drillName.Equals(defaultDrillName, StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.LogInfo($"Drill {defaultDrillName} is at default value or empty. Skipping.");
                        continue;
                    }

                    // Generate DrillData for each drill without saving to file
                    DrillData drillData = GenerateJsonForDrill(i, coordinateTable, collectDataOnly: true);
                    if (drillData != null)
                    {
                        allDrillData.Add(drillData);
                    }
                }

                if (allDrillData.Count == 0)
                {
                    MessageBox.Show("No drills have valid names to generate JSON.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Logger.LogWarning("No drills have valid names to generate JSON.");
                    return;
                }

                // Compile all drill data into WellsData
                WellsData wellsData = new WellsData
                {
                    Wells = allDrillData
                };

                // Save the compiled data to a JSON file
                SaveWellsDataToJson(wellsData);

                MessageBox.Show("JSON file generated for all drills.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("JSON file generated for all drills.");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error in HandleAllSelection: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError($"Exception in HandleAllSelection: {ex.Message}\n{ex.StackTrace}");
            }
        }
        private void HandleSingleSelection(int drillIndex)
        {
            try
            {
                if (drillIndex < 0 || drillIndex >= DrillCount)
                {
                    MessageBox.Show("Invalid drill index selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Logger.LogError($"Invalid drill index: {drillIndex}");
                    return;
                }

                // Generate JSON for the selected drill
                GenerateJsonForDrill(drillIndex);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error in HandleSingleSelection: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError($"Exception in HandleSingleSelection: {ex.Message}\n{ex.StackTrace}");
            }
        }
        private void SaveWellsDataToJson(WellsData wellsData)
        {
            try
            {
                // Serialize the WellsData object without indentation (single line)
                string jsonString = JsonConvert.SerializeObject(wellsData, Formatting.None);

                // Prompt the user to save the JSON file
                System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog
                {
                    FileName = $"{GetDrawingName()}_AllDrillsData.json",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    DefaultExt = "json",
                    AddExtension = true
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, jsonString);
                    Logger.LogInfo($"JSON file saved to {saveFileDialog.FileName}");
                }
                else
                {
                    Logger.LogInfo("User cancelled JSON save dialog.");
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error saving WellsData JSON file: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while saving the JSON file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        #region Event Handlers

        private void HeadingButton_Click(int index)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    string drillName = drillLabels[index].Text.Trim();
                    string defaultName = $"DRILL_{index + 1}";

                    Logger.LogInfo($"Initiating creation of DRILL and Heading blocks for {defaultName} with name '{drillName}'.");

                    if (string.IsNullOrWhiteSpace(drillName))
                    {
                        Logger.LogWarning($"Drill name for {defaultName} is empty. Operation aborted.");
                        MessageBox.Show($"Drill name for {defaultName} cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    Logger.LogInfo("Prompting user to select insertion point for DRILL block.");
                    Point3d insertionPointDrill;
                    PromptPointResult ppr1 = doc.Editor.GetPoint("\nSelect insertion point for DRILL block:");
                    if (ppr1.Status != PromptStatus.OK)
                    {
                        Logger.LogInfo("User canceled the selection of insertion point for DRILL block.");
                        return;
                    }
                    insertionPointDrill = ppr1.Value;
                    Logger.LogInfo($"User selected insertion point for DRILL block at {insertionPointDrill}.");

                    ObjectId drillBlockId = AutoCADHelper.InsertBlock(
                        blockName: "DRILL",
                        insertionPoint: insertionPointDrill,
                        attributes: new Dictionary<string, string> { { "DRILL", drillName } },
                        scale: 2.0
                    );
                    if (drillBlockId == ObjectId.Null)
                    {
                        Logger.LogError("Failed to insert DRILL block.");
                        MessageBox.Show("Failed to insert DRILL block.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    Logger.LogInfo($"Successfully inserted DRILL block '{drillName}' (ObjectId={drillBlockId}).");

                    string selectedHeading = headingComboBox.SelectedItem?.ToString() ?? "OTHER";
                    Logger.LogInfo($"Prompting user to select insertion point for '{selectedHeading}' block.");
                    Point3d insertionPointHeading;
                    PromptPointResult ppr2 = doc.Editor.GetPoint($"\nSelect insertion point for {selectedHeading} block:");
                    if (ppr2.Status != PromptStatus.OK)
                    {
                        Logger.LogInfo($"User canceled insertion for {selectedHeading} block.");
                        return;
                    }
                    insertionPointHeading = ppr2.Value;
                    Logger.LogInfo($"User selected insertion point for {selectedHeading} at {insertionPointHeading}.");

                    string headingBlockName = (selectedHeading == "VEREN") ? "DRILL HEADING VEREN" : "DRILL HEADING";
                    Logger.LogInfo($"Heading block name: '{headingBlockName}'.");

                    ObjectId headingBlockId = AutoCADHelper.InsertBlock(
                        blockName: headingBlockName,
                        insertionPoint: insertionPointHeading,
                        attributes: new Dictionary<string, string> { { "DRILLNAME", drillName } },
                        scale: 1.0
                    );
                    if (headingBlockId == ObjectId.Null)
                    {
                        Logger.LogError($"Failed to insert heading block '{headingBlockName}'.");
                        MessageBox.Show($"Failed to insert {headingBlockName} block.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    Logger.LogInfo($"Successfully inserted {headingBlockName} block (ObjectId={headingBlockId}).");

                    tr.Commit();
                    Logger.LogInfo($"Successfully inserted DRILL and {headingBlockName} for {defaultName}.");
                    MessageBox.Show($"Successfully inserted DRILL and {headingBlockName} blocks for {defaultName}.",
                                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }


        private void HeadingAllButton_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show(
                "Are you sure you want to insert heading blocks (and DRILL blocks) for all non-default drills?",
                "Confirm HEADING ALL", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmResult != DialogResult.Yes)
            {
                return;
            }

            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityOptions peo = new PromptEntityOptions("\nSelect the data-linked table containing 'SURFACE':");
            peo.SetRejectMessage("\nOnly table entities are allowed.");
            peo.AddAllowedClass(typeof(Table), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nTable selection cancelled.");
                return;
            }

            Table surfaceTable;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                surfaceTable = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                tr.Commit();
            }

            if (surfaceTable == null)
            {
                MessageBox.Show("Selected entity is not a valid table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<(int row, int col)> surfaceCells = new List<(int row, int col)>();
            for (int r = 0; r < surfaceTable.Rows.Count; r++)
            {
                for (int c = 0; c < surfaceTable.Columns.Count; c++)
                {
                    string cellValue = surfaceTable.Cells[r, c].TextString.Trim().ToUpper();
                    if (cellValue.Contains("SURFACE"))
                    {
                        surfaceCells.Add((r, c));
                    }
                }
            }

            if (surfaceCells.Count == 0)
            {
                MessageBox.Show("No 'SURFACE' cells found in the selected table.", "No Data",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<string> nonDefaultDrills = new List<string>();
            for (int i = 0; i < DrillCount; i++)
            {
                string drillName = drillLabels[i].Text.Trim();
                string defaultName = $"DRILL_{i + 1}";
                if (!string.IsNullOrWhiteSpace(drillName) &&
                    !drillName.Equals(defaultName, StringComparison.OrdinalIgnoreCase))
                {
                    nonDefaultDrills.Add(drillName);
                }
            }

            if (nonDefaultDrills.Count == 0)
            {
                MessageBox.Show("No non-default drills to insert blocks for.", "No Data",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (nonDefaultDrills.Count > surfaceCells.Count)
            {
                MessageBox.Show("Not enough SURFACE cells for all non-default drills.",
                                "Data Mismatch", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string headingBlockName = headingComboBox.SelectedItem.ToString() == "VEREN"
                                       ? "DRILL HEADING VEREN"
                                       : "DRILL HEADING";

            try
            {
                using (Transaction tr = AutoCADHelper.StartTransaction())
                {
                    for (int i = 0; i < nonDefaultDrills.Count; i++)
                    {
                        string drillName = nonDefaultDrills[i];
                        var (surfaceRow, surfaceCol) = surfaceCells[i];

                        Point3d nwCorner = GetCellNWCorner(surfaceTable, surfaceRow, surfaceCol);

                        ObjectId headingBlockId = AutoCADHelper.InsertBlock(
                            blockName: headingBlockName,
                            insertionPoint: nwCorner,
                            attributes: new Dictionary<string, string> { { "DRILLNAME", drillName } },
                            scale: 1.0
                        );
                        if (headingBlockId == ObjectId.Null)
                        {
                            MessageBox.Show($"Failed to insert {headingBlockName} for {drillName}.",
                                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                        }

                        Point3d drillPoint = new Point3d(nwCorner.X - 50.0, nwCorner.Y, nwCorner.Z);
                        ObjectId drillBlockId = AutoCADHelper.InsertBlock(
                            blockName: "DRILL",
                            insertionPoint: drillPoint,
                            attributes: new Dictionary<string, string> { { "DRILL", drillName } },
                            scale: 2.0
                        );
                        if (drillBlockId == ObjectId.Null)
                        {
                            MessageBox.Show($"Failed to insert DRILL block for {drillName}.",
                                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                        }
                    }
                    tr.Commit();
                }
                MessageBox.Show("Successfully created DRILL + HEADING blocks for all non-default drills.",
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Exception in HeadingAllButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while inserting heading blocks: {ex.Message}",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CompleteCordsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // STEP 0: Extract grid labels from layer "Z-DRILL-POINT" and export to C:\CORDS\CORDS.csv
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;
                var gridData = new List<(string Label, double Northing, double Easting)>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    foreach (ObjectId id in ms)
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null && ent.Layer.Equals("Z-DRILL-POINT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (ent is DBText dbText)
                            {
                                string textValue = dbText.TextString.Trim();
                                if (IsGridLabel(textValue))
                                {
                                    gridData.Add((textValue, dbText.Position.Y, dbText.Position.X));
                                }
                            }
                            else if (ent is MText mText)
                            {
                                string textValue = mText.Contents.Trim();
                                if (IsGridLabel(textValue))
                                {
                                    gridData.Add((textValue, mText.Location.Y, mText.Location.X));
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                gridData = gridData.OrderBy(x => x.Label).ToList();

                // STEP 1: Write CSV file
                string cordsDir = @"C:\CORDS";
                Directory.CreateDirectory(cordsDir);
                string csvPath = Path.Combine(cordsDir, "cords.csv");
                using (StreamWriter sw = new StreamWriter(csvPath, false))
                {
                    sw.WriteLine("Label,Northing,Easting");
                    foreach (var pt in gridData)
                    {
                        sw.WriteLine($"{pt.Label},{pt.Northing},{pt.Easting}");
                    }
                }

                MessageBox.Show("DONT TOUCH, WAIT FOR INSTRUCTION", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // STEP 2: Determine parameter value.
                string headingValue = headingComboBox.SelectedItem?.ToString() ?? "OTHER";
                string paramValue = (headingValue == "VEREN") ? "HEEL" : "ICP";

                // STEP 3: Launch the Python EXE (cords.exe). With the updated Python script it will close once the Excel file is generated.
                string processExe = @"C:\AUTOCAD-SETUP\Lisp_2000\Drill Properties\cords.exe";
                if (File.Exists(processExe))
                {
                    Logger.LogInfo($"Attempting to run exe: {processExe} with arguments: \"{csvPath}\" \"{paramValue}\"");
                    ProcessStartInfo psi = new ProcessStartInfo(processExe, $"\"{csvPath}\" \"{paramValue}\"")
                    {
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };
                    // Ensure output encoding matches Python's cp1252 fallback
                    psi.StandardOutputEncoding = Encoding.GetEncoding(1252);
                    psi.StandardErrorEncoding = Encoding.GetEncoding(1252);

                    using (Process proc = Process.Start(psi))
                    {
                        // Wait up to 180 seconds for the Python process to finish.
                        if (!proc.WaitForExit(180000))
                        {
                            try { proc.Kill(); } catch { }
                            Logger.LogError("cords.exe did not exit within 180 seconds.");
                            MessageBox.Show("The cords.exe process did not exit in time and was terminated.",
                                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        string output = proc.StandardOutput.ReadToEnd();
                        string errorOutput = proc.StandardError.ReadToEnd();
                        if (!string.IsNullOrEmpty(output))
                            Logger.LogInfo($"cords.exe output: {output}");
                        if (!string.IsNullOrEmpty(errorOutput))
                            Logger.LogError($"cords.exe error: {errorOutput}");
                        if (proc.ExitCode != 0)
                        {
                            Logger.LogError($"cords.exe exited with code {proc.ExitCode}");
                            MessageBox.Show($"cords.exe exited with code {proc.ExitCode}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    Logger.LogInfo($"cords.exe executed successfully. Param = {paramValue}");
                }
                else
                {
                    Logger.LogWarning($"Executable not found at {processExe}. Skipping execution.");
                }

                // STEP 4: Confirm the Excel file exists.
                string excelFilePath = Path.Combine(cordsDir, "ExportedCoordsFormatted.xlsx");
                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show("The file ExportedCoordsFormatted.xlsx was not found in C:\\CORDS.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // STEP 5: Read and process the Excel file.
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                string[,] tableData;
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null)
                    {
                        MessageBox.Show("The Excel file is empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    int startRow = ws.Dimension.Start.Row;
                    int startCol = ws.Dimension.Start.Column;
                    int endRow = ws.Dimension.End.Row;
                    int endCol = ws.Dimension.End.Column;

                    int lastRow = endRow;
                    for (int r = endRow; r >= startRow; r--)
                    {
                        bool rowBlank = true;
                        for (int c = startCol; c <= endCol; c++)
                        {
                            if (!string.IsNullOrWhiteSpace(ws.Cells[r, c].Text))
                            {
                                rowBlank = false;
                                break;
                            }
                        }
                        if (!rowBlank)
                        {
                            lastRow = r;
                            break;
                        }
                    }

                    int rows = lastRow - startRow + 1;
                    int cols = endCol - startCol + 1;
                    tableData = new string[rows, cols];
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            tableData[r, c] = ws.Cells[r + startRow, c + startCol].Text;
                        }
                    }
                }
                Logger.LogInfo($"Read {tableData.GetLength(0)} rows and {tableData.GetLength(1)} columns from {excelFilePath}.");

                // STEP 6: Adjust cell values based on client selection.
                if (headingComboBox.SelectedItem?.ToString() == "VEREN")
                {
                    for (int r = 0; r < tableData.GetLength(0); r++)
                    {
                        for (int c = 0; c < tableData.GetLength(1); c++)
                        {
                            if (!string.IsNullOrEmpty(tableData[r, c]) && tableData[r, c].Contains("ICP"))
                            {
                                tableData[r, c] = tableData[r, c].Replace("ICP", "HEEL");
                            }
                        }
                    }
                }
                else if (headingComboBox.SelectedItem?.ToString() == "OTHER")
                {
                    for (int r = 0; r < tableData.GetLength(0); r++)
                    {
                        for (int c = 0; c < tableData.GetLength(1); c++)
                        {
                            if (!string.IsNullOrEmpty(tableData[r, c]) && tableData[r, c].Contains("HEEL"))
                            {
                                tableData[r, c] = tableData[r, c].Replace("HEEL", "ICP");
                            }
                        }
                    }
                }

                // STEP 7: Prompt user for insertion point.
                MessageBox.Show("BACK TO CAD, PICK A INSERTION POINT", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                doc = AcApplication.DocumentManager.MdiActiveDocument;
                ed = doc.Editor;
                PromptPointResult ppr = ed.GetPoint("\nSelect insertion point for the coordinate table:");
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nInsertion point selection cancelled.");
                    return;
                }
                Point3d insertionPt = ppr.Value;

                // STEP 8: Insert and format the table.
                using (DocumentLock docLock = doc.LockDocument())
                {
                    db = doc.Database;
                    EnsureLayer(db, "CG-NOTES");
                    InsertAndFormatTable(insertionPt, tableData, "induction Bend");
                }
                MessageBox.Show("Coordinate table created successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("COMPLETE CORDS: Coordinate table created successfully.");

                // STEP 9: Insert headings.
                HeadingAllButton_Click(null, EventArgs.Empty);
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in COMPLETE CORDS: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error in COMPLETE CORDS: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Helper method to determine if a text string is a grid label in the range A1–L12.
        /// </summary>
        private bool IsGridLabel(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            // Allow any uppercase letter from A to Z (modify if you need a different range)
            char letter = text[0];
            if (letter < 'A' || letter > 'Z')
                return false;

            // The rest of the text should represent a number between 1 and 150.
            string numberPart = text.Substring(1);
            if (int.TryParse(numberPart, out int num))
            {
                return num >= 1 && num <= 150;
            }
            return false;
        }

        /// <summary>
        /// Helper method to determine if a text string is a grid label in the range A1–L12.
        /// </summary>

        private void SwapButton_Click(object sender, EventArgs e)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            int index1 = swapComboBox1.SelectedIndex;
            int index2 = swapComboBox2.SelectedIndex;

            if (index1 == index2)
            {
                MessageBox.Show("Please pick two different drills to swap.");
                return;
            }

            string oldText1 = drillTextBoxes[index1].Text.Trim();
            string oldText2 = drillTextBoxes[index2].Text.Trim();
            string oldLabel1 = drillLabels[index1].Text.Trim();
            string oldLabel2 = drillLabels[index2].Text.Trim();

            string newText1 = oldText2;
            string newText2 = oldText1;
            string newLabel1 = oldLabel2;
            string newLabel2 = oldLabel1;

            drillTextBoxes[index1].Text = newText1;
            drillTextBoxes[index2].Text = newText2;
            drillLabels[index1].Text = newLabel1;
            drillLabels[index2].Text = newLabel2;

            using (DocumentLock docLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt == null) return;
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null) continue;
                        foreach (ObjectId entId in btr)
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                            if (ent is BlockReference blockRef)
                            {
                                string tag1 = $"DRILL_{index1 + 1}";
                                string tag2 = $"DRILL_{index2 + 1}";
                                SwapDrillAttribute(blockRef, tag1, newText1, tr);
                                SwapDrillAttribute(blockRef, tag2, newText2, tr);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            SaveToJson();

            MessageBox.Show($"Swapped {oldText1} <-> {oldText2}", "Swap Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SetAllButton_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show(
                "Are you sure you want to set DRILLNAME for all drills?",
                "Confirm SET ALL", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmResult != DialogResult.Yes)
            {
                Logger.LogInfo("Set All operation canceled by the user.");
                return;
            }

            List<string> changes = new List<string>();
            int updatedCount = 0;
            for (int i = 0; i < DrillCount; i++)
            {
                string drillName = drillTextBoxes[i].Text.Trim();
                string defaultName = GetDefaultDrillName(i);
                bool isDefault = string.Equals(drillName, defaultName, StringComparison.OrdinalIgnoreCase);
                if (!isDefault || i == 0)
                {
                    string oldName = drillLabels[i].Text.Trim();
                    if (!string.Equals(oldName, drillName, StringComparison.OrdinalIgnoreCase))
                    {
                        changes.Add($"{defaultName}: '{oldName}' -> '{drillName}'");
                    }
                    SetDrill(i, false);
                    updatedCount++;
                }
            }

            SaveToJson();
            Logger.LogInfo($"Set All operation completed. Total drills updated: {updatedCount}.");
            string summary = changes.Count > 0 ? string.Join("\n", changes) : "No changes were necessary.";
            MessageBox.Show(summary, "Set All", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ResetDrill(int index)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            string defaultName = $"DRILL_{index + 1}";
            // DRILL_1 remains DRILL_1, others become blank
            string newValue = (index == 0) ? defaultName : string.Empty;

            using (DocumentLock docLock = doc.LockDocument())
            {
                try
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = null;
                        try
                        {
                            bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        }
                        catch { return; }

                        if (bt == null)
                            return;

                        int updatedAttributes = 0;

                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = null;
                            try
                            {
                                btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                            }
                            catch { continue; }
                            if (btr == null || btr.IsErased)
                                continue;

                            foreach (ObjectId entId in btr)
                            {
                                Entity ent = null;
                                try
                                {
                                    ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                                }
                                catch { continue; }
                                if (ent == null || ent.IsErased)
                                    continue;

                                if (ent is BlockReference blockRef)
                                {
                                    foreach (ObjectId attId in blockRef.AttributeCollection)
                                    {
                                        AttributeReference attRef = null;
                                        try
                                        {
                                            attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                                        }
                                        catch { continue; }
                                        if (attRef == null || attRef.IsErased)
                                            continue;

                                        if (attRef.Tag.Equals(defaultName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            attRef.TextString = newValue;
                                            updatedAttributes++;
                                        }
                                    }
                                }

                                if (ent is MText mText)
                                {
                                    string oldName = drillTextBoxes[index].Text.Trim();
                                    if (!string.IsNullOrEmpty(oldName) && mText.Contents.Contains(oldName))
                                    {
                                        mText.Contents = mText.Contents.Replace(oldName, defaultName);
                                        updatedAttributes++;
                                    }
                                }
                            }
                        }
                        tr.Commit();
                        Logger.LogInfo($"ResetDrill => updated {updatedAttributes} attributes.");
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.LogError($"Error in ResetDrill: {ex.Message}\n{ex.StackTrace}");
                    MessageBox.Show($"Error: {ex.Message}", "Reset Drill", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Update the UI
            drillTextBoxes[index].Text = newValue;
            drillLabels[index].Text = defaultName;
            SaveToJson();

            MessageBox.Show($"Successfully reset {defaultName}.", "Reset Drill",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void ResetAllButton_Click(object sender, EventArgs e)
        {
            Logger.LogInfo("Initiating Reset All operation.");
            int totalReset = 0;

            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            try
            {
                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = null;
                        try
                        {
                            bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        }
                        catch { return; }

                        if (bt == null)
                            return;

                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = null;
                            try
                            {
                                btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                            }
                            catch { continue; }
                            if (btr == null || btr.IsErased)
                                continue;

                            // Upgrade btr for writing
                            btr.UpgradeOpen();

                            foreach (ObjectId entId in btr)
                            {
                                Entity ent = null;
                                try
                                {
                                    ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                                }
                                catch { continue; }
                                if (ent == null || ent.IsErased)
                                    continue;

                                if (ent is BlockReference blockRef)
                                {
                                    // For each drill, reset the attribute
                                    for (int i = 0; i < DrillCount; i++)
                                    {
                                        string defaultName = $"DRILL_{i + 1}";
                                        string newValue = (i == 0) ? defaultName : string.Empty;

                                        foreach (ObjectId attId in blockRef.AttributeCollection)
                                        {
                                            AttributeReference attRef = null;
                                            try
                                            {
                                                attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                                            }
                                            catch { continue; }
                                            if (attRef == null || attRef.IsErased)
                                                continue;

                                            if (attRef.Tag.Equals(defaultName, StringComparison.OrdinalIgnoreCase))
                                            {
                                                Logger.LogInfo($"Resetting '{attRef.Tag}' from '{attRef.TextString}' to '{newValue}'.");
                                                attRef.TextString = newValue;
                                                totalReset++;
                                            }
                                        }
                                    }
                                }

                                if (ent is MText mText)
                                {
                                    for (int i = 0; i < DrillCount; i++)
                                    {
                                        string defaultName = $"DRILL_{i + 1}";
                                        string currentText = drillTextBoxes[i].Text.Trim();
                                        string newValue = (i == 0) ? defaultName : string.Empty;

                                        if (!string.IsNullOrEmpty(currentText) && mText.Contents.Contains(currentText))
                                        {
                                            Logger.LogInfo($"Resetting MText from '{mText.Contents}' to '{newValue}' (Handle={mText.Handle}).");
                                            mText.Contents = mText.Contents.Replace(currentText, newValue);
                                            totalReset++;
                                        }
                                    }
                                }
                            }
                        }
                        tr.Commit();
                    }
                }

                // Update UI after resetting
                for (int i = 0; i < DrillCount; i++)
                {
                    string defaultName = $"DRILL_{i + 1}";
                    drillTextBoxes[i].Text = (i == 0) ? defaultName : string.Empty;
                    drillLabels[i].Text = defaultName;
                }

                SaveToJson();

                Logger.LogInfo($"Reset All completed. {totalReset} attributes reset.");
                MessageBox.Show($"Successfully reset all drills. Total attributes reset: {totalReset}.",
                                "Reset All", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error during Reset All: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred during the Reset All operation: {ex.Message}",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateFromAttributesButton_Click(object sender, EventArgs e)
        {
            bool anyUpdatePerformed = false;
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                PromptSelectionResult selResult = ed.GetSelection();
                if (selResult.Status != PromptStatus.OK)
                {
                    MessageBox.Show("No objects selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Dictionary<int, bool> updatedDrillAttributes = new Dictionary<int, bool>();
                using (Transaction tr = AutoCADHelper.StartTransaction())
                {
                    SelectionSet selSet = selResult.Value;
                    foreach (SelectedObject selObj in selSet)
                    {
                        if (selObj != null)
                        {
                            Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent is BlockReference blockRef)
                            {
                                foreach (ObjectId attId in blockRef.AttributeCollection)
                                {
                                    AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                    if (attRef != null)
                                    {
                                        for (int i = 1; i <= DrillCount; i++)
                                        {
                                            string drillTag = $"DRILL_{i}";
                                            if (attRef.Tag.Equals(drillTag, StringComparison.OrdinalIgnoreCase))
                                            {
                                                if (!string.IsNullOrWhiteSpace(attRef.TextString))
                                                {
                                                    drillTextBoxes[i - 1].Text = attRef.TextString;
                                                    drillLabels[i - 1].Text = attRef.TextString;
                                                    anyUpdatePerformed = true;
                                                    updatedDrillAttributes[i] = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    for (int i = 1; i <= DrillCount; i++)
                    {
                        if (!updatedDrillAttributes.ContainsKey(i))
                        {
                            string defaultName = $"DRILL_{i}";
                            drillTextBoxes[i - 1].Text = defaultName;
                            drillLabels[i - 1].Text = defaultName;
                            anyUpdatePerformed = true;
                        }
                    }
                    tr.Commit();
                }

                if (anyUpdatePerformed)
                {
                    MessageBox.Show("Form fields have been updated from selected block attributes.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    SaveToJson();
                }
                else
                {
                    MessageBox.Show("No matching DRILL_# attributes found to update form fields.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"An error occurred while updating from block attributes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void CreateTableButton_Click(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                    Title = "Select Excel File"
                };
                if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;
                string excelFilePath = ofd.FileName;

                string range = ShowInputDialog("Enter cell range (e.g. A1:L15):", "Cell Range", "A1:L15");
                if (string.IsNullOrWhiteSpace(range))
                    return;

                string[,] excelData = ReadExcelData(excelFilePath, range);
                if (excelData == null)
                {
                    MessageBox.Show("Failed to read Excel data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                PromptPointResult ppr = ed.GetPoint("\nSelect insertion point for the table:");
                if (ppr.Status != PromptStatus.OK)
                    return;
                Point3d insertionPoint = ppr.Value;

                string tableStyleName = "induction Bend";

                using (DocumentLock docLock = doc.LockDocument())
                {
                    // Ensure that the "CG-NOTES" layer exists.
                    EnsureLayer(doc.Database, "CG-NOTES");
                    // Now, create the table. InsertAndFormatTable will set the table's layer to "CG-NOTES".
                    InsertAndFormatTable(insertionPoint, excelData, tableStyleName);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error creating table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateXlsButton_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                PromptEntityOptions peo = new PromptEntityOptions("\nSelect the table to export to XLS:");
                peo.SetRejectMessage("\nOnly table entities are allowed.");
                peo.AddAllowedClass(typeof(Table), true);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nXLS save cancelled.");
                    return;
                }

                Table selectedTable;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    selectedTable = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                    tr.Commit();
                }
                if (selectedTable == null)
                {
                    MessageBox.Show("The selected entity is not a valid table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    Title = "Save XLS File",
                    FileName = "ExportedTable.xlsx"
                };

                if (sfd.ShowDialog() != DialogResult.OK)
                {
                    ed.WriteMessage("\nXLS save cancelled.");
                    return;
                }
                string excelFilePath = sfd.FileName;

                int rows = selectedTable.Rows.Count;
                int cols = selectedTable.Columns.Count;

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("ExportedTable");

                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            string cellValue = selectedTable.Cells[r, c].TextString.Trim();
                            ws.Cells[r + 1, c + 1].Value = cellValue;
                        }
                    }

                    if (cols >= 3)
                    {
                        ws.Column(1).Width = 15;
                        ws.Column(2).Width = 12;
                        ws.Column(3).Width = 12;
                        ws.Column(2).Style.Numberformat.Format = "0.00";
                        ws.Column(3).Style.Numberformat.Format = "0.00";
                    }

                    package.SaveAs(new FileInfo(excelFilePath));
                }

                MessageBox.Show($"XLS file created successfully at:\n{excelFilePath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"An error occurred while creating XLS:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WellCornersButton_Click(object sender, EventArgs e)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polygon (polyline):");
                peo.SetRejectMessage("\nOnly polylines are allowed.");
                peo.AddAllowedClass(typeof(Polyline), true);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nSelection cancelled.");
                    return;
                }

                List<Point2d> vertices = new List<Point2d>();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Polyline poly = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
                    if (poly == null)
                    {
                        MessageBox.Show("Selected entity is not a valid polyline.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    int numVerts = poly.NumberOfVertices;
                    for (int i = 0; i < numVerts; i++)
                    {
                        vertices.Add(poly.GetPoint2dAt(i));
                    }
                    tr.Commit();
                }
                if (vertices.Count == 0)
                {
                    MessageBox.Show("No vertices found in the selected polygon.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PromptPointResult ppr = ed.GetPoint("\nSelect insertion point for the WELL CORNERS table:");
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nInsertion point cancelled.");
                    return;
                }
                Point3d insertionPoint = ppr.Value;

                using (DocumentLock docLock = doc.LockDocument())
                {
                    EnsureLayer(db, "CG-NOTES");
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        int rows = vertices.Count;
                        int cols = 2;
                        Table table = new Table();
                        table.TableStyle = GetTableStyleId(db, "induction Bend", tr);
                        table.SetSize(rows, cols);
                        table.Position = insertionPoint;
                        table.Layer = "CG-NOTES";

                        for (int c = 0; c < cols; c++)
                        {
                            table.Columns[c].Width = 89.41;
                        }

                        for (int r = 0; r < rows; r++)
                        {
                            table.Rows[r].Height = 25.0;
                            Point2d vert = vertices[r];
                            table.Cells[r, 0].TextString = vert.Y.ToString("F2");
                            table.Cells[r, 1].TextString = vert.X.ToString("F2");
                        }

                        btr.AppendEntity(table);
                        tr.AddNewlyCreatedDBObject(table, true);
                        table.GenerateLayout();
                        UnmergeAllCells(table);
                        table.GenerateLayout();
                        tr.Commit();
                    }
                    MessageBox.Show("WELL CORNERS table created successfully on CG-NOTES, with unmerged cells.",
                                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error creating WELL CORNERS table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Returns the NW (top-left) corner of a cell in the table.
        /// </summary>
        private Point3d GetCellNWCorner(Table table, int row, int col)
        {
            double baseX = table.Position.X;
            double baseY = table.Position.Y;
            double xOffset = baseX;
            for (int i = 0; i < col; i++)
            {
                xOffset += table.Columns[i].Width;
            }
            double yOffset = baseY;
            for (int i = 0; i < row; i++)
            {
                yOffset -= table.Rows[i].Height;
            }
            return new Point3d(xOffset, yOffset, 0.0);
        }
        #endregion

        #region Utility Functions



        private string GetJsonFilePath()
        {
            var acDoc = AcApplication.DocumentManager.MdiActiveDocument;
            if (acDoc == null || string.IsNullOrEmpty(acDoc.Name))
            {
                MessageBox.Show("No active AutoCAD document found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new InvalidOperationException("No active AutoCAD document found.");
            }
            string dwgDirectory = Path.GetDirectoryName(acDoc.Name);
            string drawingName = Path.GetFileNameWithoutExtension(acDoc.Name);
            drawingName = drawingName.TrimEnd('-');
            string[] nameParts = drawingName.Split('-');
            string prefix = nameParts.Length >= 2 ? $"{nameParts[0]}-{nameParts[1]}" : drawingName;
            string jsonFilePath = Path.Combine(dwgDirectory, $"{prefix}.json");
            return jsonFilePath;
        }

        private string GetDrawingNameWithoutExtension()
        {
            var acDoc = AcApplication.DocumentManager.MdiActiveDocument;
            if (acDoc == null || string.IsNullOrEmpty(acDoc.Name))
            {
                throw new InvalidOperationException("No active AutoCAD document found.");
            }
            string drawingName = Path.GetFileNameWithoutExtension(acDoc.Name);
            return drawingName;
        }

        private string ShowInputDialog(string prompt, string title, string defaultValue = "")
        {
            using (Form promptForm = new Form())
            {
                promptForm.Width = 300;
                promptForm.Height = 150;
                promptForm.Text = title;
                promptForm.BackColor = System.Drawing.Color.Black;
                promptForm.ForeColor = System.Drawing.Color.White;
                Label textLabel = new Label() { Left = 10, Top = 20, Text = prompt, AutoSize = true };
                textLabel.ForeColor = System.Drawing.Color.White;
                TextBox inputBox = new TextBox() { Left = 10, Top = 50, Width = 260 };
                inputBox.Text = defaultValue;
                Button okButton = new Button() { Text = "OK", Left = 200, Width = 70, Top = 80 };
                okButton.DialogResult = DialogResult.OK;
                promptForm.Controls.Add(textLabel);
                promptForm.Controls.Add(inputBox);
                promptForm.Controls.Add(okButton);
                promptForm.AcceptButton = okButton;
                return promptForm.ShowDialog() == DialogResult.OK ? inputBox.Text : null;
            }
        }

        private void InitializeDefaultDrills()
        {
            for (int i = 0; i < DrillCount; i++)
            {
                string defaultName = $"DRILL_{i + 1}";
                drillTextBoxes[i].Text = defaultName;
                drillLabels[i].Text = defaultName;
            }
        }

        private List<string> ExtractTableValues(Table table)
        {
            List<string> tableValues = new List<string>();
            for (int row = 0; row < table.Rows.Count; row++)
            {
                string columnAValue = table.Cells[row, 0].TextString.Trim();
                string normalizedColumnAValue = columnAValue.ToUpper().Replace(" ", "");
                if (normalizedColumnAValue.Contains("BOTTOMHOLE"))
                {
                    string columnBValue = table.Cells[row, 1].TextString.Trim();
                    tableValues.Add(columnBValue);
                    Logger.LogInfo($"Found 'BOTTOM HOLE' at row {row}. Column B value: '{columnBValue}'");
                }
                else
                {
                    Logger.LogInfo($"Row {row}: Column A value '{columnAValue}' does not match 'BOTTOM HOLE'");
                }
            }
            if (tableValues.Count == 0)
            {
                MessageBox.Show("No 'BOTTOM HOLE' entries found in the selected table.", "Check Results", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Logger.LogWarning("No 'BOTTOM HOLE' entries found in the selected table.");
            }
            return tableValues;
        }



        private void CompareDrillNamesWithTableAndBlocks(List<string> tableValues, List<BlockAttributeData> blockDataList)
        {
            try
            {
                // Sort the blocks based on Y-coordinate (descending)
                blockDataList.Sort((a, b) => b.YCoordinate.CompareTo(a.YCoordinate));

                int comparisons = Math.Min(Math.Min(tableValues.Count, DrillCount), blockDataList.Count);
                List<string> discrepancies = new List<string>();
                List<string> reportLines = new List<string>();

                for (int i = 0; i < comparisons; i++)
                {
                    // From table
                    string tableValue = tableValues[i];
                    string normalizedTableValue = NormalizeTableValue(tableValue);

                    // From application
                    string drillName = drillTextBoxes[i].Text.Trim();
                    string normalizedDrillName = NormalizeDrillName(drillName);

                    // From block
                    BlockAttributeData blockData = blockDataList[i];
                    string blockDrillName = blockData.DrillName;
                    string normalizedBlockDrillName = NormalizeDrillName(blockDrillName);

                    bool discrepancyFound = false;
                    List<string> drillDiscrepancies = new List<string>();

                    // Compare normalized values
                    if (string.IsNullOrEmpty(normalizedDrillName))
                    {
                        drillDiscrepancies.Add($"Unable to extract numeric part from drill name '{drillName}'.");
                        discrepancyFound = true;
                    }
                    if (string.IsNullOrEmpty(normalizedTableValue))
                    {
                        drillDiscrepancies.Add($"Unable to extract numeric part from table value '{tableValue}'.");
                        discrepancyFound = true;
                    }
                    if (string.IsNullOrEmpty(normalizedBlockDrillName))
                    {
                        drillDiscrepancies.Add($"Unable to extract numeric part from block DRILLNAME '{blockDrillName}'.");
                        discrepancyFound = true;
                    }

                    if (!discrepancyFound)
                    {
                        // Compare drill name with table value
                        if (!string.Equals(normalizedDrillName, normalizedTableValue, StringComparison.OrdinalIgnoreCase))
                        {
                            drillDiscrepancies.Add($"Drill name does not match table value.");
                            discrepancyFound = true;
                        }

                        // Compare block DRILLNAME with drill name
                        if (!string.Equals(normalizedBlockDrillName, normalizedDrillName, StringComparison.OrdinalIgnoreCase))
                        {
                            drillDiscrepancies.Add($"Block DRILLNAME does not match drill name.");
                            discrepancyFound = true;
                        }

                        // Compare block DRILLNAME with table value
                        if (!string.Equals(normalizedBlockDrillName, normalizedTableValue, StringComparison.OrdinalIgnoreCase))
                        {
                            drillDiscrepancies.Add($"Block DRILLNAME does not match table value.");
                            discrepancyFound = true;
                        }
                    }

                    // Build the report
                    reportLines.Add($"DRILL_{i + 1} NAME: {drillName}");
                    reportLines.Add($"TABLE RESULT: {tableValue}");
                    reportLines.Add($"BLOCK DRILLNAME: {blockDrillName}");
                    reportLines.Add($"Normalized Drill Name: {normalizedDrillName}");
                    reportLines.Add($"Normalized Table Value: {normalizedTableValue}");
                    reportLines.Add($"Normalized Block DrillName: {normalizedBlockDrillName}");

                    string status = discrepancyFound ? "FAIL" : "PASS";
                    reportLines.Add($"STATUS: {status}");

                    // Include discrepancies in the report if any
                    if (discrepancyFound)
                    {
                        reportLines.Add("Discrepancies:");
                        foreach (var disc in drillDiscrepancies)
                        {
                            reportLines.Add($"- {disc}");
                        }
                        discrepancies.Add($"DRILL_{i + 1}: {string.Join("; ", drillDiscrepancies)}");
                        HighlightTextbox(drillTextBoxes[i], true);
                    }
                    else
                    {
                        HighlightTextbox(drillTextBoxes[i], false);
                    }

                    reportLines.Add(""); // Blank line for spacing

                    // Logging
                    Logger.LogInfo($"Comparing DRILL_{i + 1}:");
                    Logger.LogInfo($"Original Drill Name: '{drillName}'");
                    Logger.LogInfo($"Original Table Value: '{tableValue}'");
                    Logger.LogInfo($"Original Block DrillName: '{blockDrillName}'");
                    Logger.LogInfo($"Normalized Drill Name: '{normalizedDrillName}'");
                    Logger.LogInfo($"Normalized Table Value: '{normalizedTableValue}'");
                    Logger.LogInfo($"Normalized Block DrillName: '{normalizedBlockDrillName}'");
                    Logger.LogInfo($"Status: {status}");
                    if (discrepancyFound)
                    {
                        Logger.LogInfo($"Discrepancies: {string.Join("; ", drillDiscrepancies)}");
                    }
                }

                // Write the report to a .txt file
                string reportFilePath = GetReportFilePath();
                try
                {
                    File.WriteAllLines(reportFilePath, reportLines);
                    Logger.LogInfo($"Report successfully written to {reportFilePath}");
                }
                catch (System.Exception ex)
                {
                    Logger.LogError($"Error writing report file: {ex.Message}");
                    MessageBox.Show($"An error occurred while writing the report file: {ex.Message}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (discrepancies.Count > 0)
                {
                    // Report discrepancies
                    string message = "Discrepancies found:\n" + string.Join("\n", discrepancies);
                    MessageBox.Show($"{message}\n\nDetailed report saved at:\n{reportFilePath}", "Check Results", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show($"All drill names match the table values and block DRILLNAME attributes.\n\nDetailed report saved at:\n{reportFilePath}", "Check Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in CompareDrillNamesWithTableAndBlocks: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred during comparison: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string NormalizeDrillName(string drillName)
        {
            Regex regex = new Regex(@"(\d{1,2}-\d{1,2}-\d{1,3}-\d{1,2})");
            Match match = regex.Match(drillName);
            if (match.Success)
            {
                string numericPart = match.Value;
                string[] parts = numericPart.Split('-');
                for (int i = 0; i < parts.Length; i++)
                {
                    parts[i] = parts[i].TrimStart('0');
                }
                string normalized = string.Join("-", parts);
                Logger.LogInfo($"Normalized drill name '{drillName}' to '{normalized}'");
                return normalized;
            }
            else
            {
                Logger.LogWarning($"Drill name '{drillName}' does not contain the expected numeric pattern.");
                return "";
            }
        }

        private string NormalizeTableValue(string tableValue)
        {
            tableValue = Regex.Replace(tableValue, @"\{.*?;", "");
            tableValue = tableValue.Replace("}", "");
            tableValue = tableValue.ToUpper().Replace(" ", "");
            tableValue = Regex.Replace(tableValue, "W\\d+", "", RegexOptions.IgnoreCase);
            string[] parts = tableValue.Split('-');
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].TrimStart('0');
            }
            string normalized = string.Join("-", parts);
            Logger.LogInfo($"Normalized table value '{tableValue}' to '{normalized}'");
            return normalized;
        }


        private string GetAttributeValue(BlockReference blockRef, string attributeTag, Transaction tr)
        {
            try
            {
                foreach (ObjectId attId in blockRef.AttributeCollection)
                {
                    AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                    if (attRef != null && attRef.Tag.Equals(attributeTag, StringComparison.OrdinalIgnoreCase))
                    {
                        Logger.LogInfo($"Extracted attribute '{attributeTag}': '{attRef.TextString}'");
                        return attRef.TextString.Trim();
                    }
                }
                Logger.LogWarning($"Attribute '{attributeTag}' not found in block '{blockRef.Name}'.");
                return string.Empty;
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetAttributeValue: {ex.Message}\n{ex.StackTrace}");
                return string.Empty;
            }
        }

        private void EnsureLayer(Database db, string layerName)
        {
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                if (!lt.Has(layerName))
                {
                    using (LayerTableRecord ltr = new LayerTableRecord())
                    {
                        ltr.Name = layerName;
                        lt.UpgradeOpen();
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }
                }
                tr.Commit();
            }
        }

        private ObjectId GetTableStyleId(Database db, string styleName, Transaction tr)
        {
            DBDictionary tableStyleDict = (DBDictionary)tr.GetObject(db.TableStyleDictionaryId, OpenMode.ForRead);
            foreach (var entry in tableStyleDict)
            {
                if (entry.Key.Equals(styleName, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }
            return ObjectId.Null;
        }

        private void UnmergeAllCells(Table table)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var range = table.Cells[r, c].GetMergeRange();
                    if (range != null && range.TopRow == r && range.LeftColumn == c)
                    {
                        table.UnmergeCells(range);
                    }
                }
            }
        }

        private void RemoveAllCellBorders(Table table)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    Cell cell = table.Cells[r, c];
                    cell.Borders.Top.IsVisible = false;
                    cell.Borders.Bottom.IsVisible = false;
                    cell.Borders.Left.IsVisible = false;
                    cell.Borders.Right.IsVisible = false;
                }
            }
        }

        private void AddBordersForDataCells(Table table)
        {
            AColor borderColor = AColor.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7);
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    Cell cell = table.Cells[r, c];
                    string val = cell.TextString;
                    bool hasData = !string.IsNullOrWhiteSpace(val);
                    if (hasData)
                    {
                        SetCellBorders(cell, true, borderColor);
                    }
                    else
                    {
                        bool topHasData = r > 0 && !string.IsNullOrWhiteSpace(table.Cells[r - 1, c].TextString);
                        bool bottomHasData = (r < table.Rows.Count - 1) && !string.IsNullOrWhiteSpace(table.Cells[r + 1, c].TextString);
                        bool leftHasData = c > 0 && !string.IsNullOrWhiteSpace(table.Cells[r, c - 1].TextString);
                        bool rightHasData = (c < table.Columns.Count - 1) && !string.IsNullOrWhiteSpace(table.Cells[r, c + 1].TextString);
                        cell.Borders.Top.IsVisible = topHasData;
                        cell.Borders.Bottom.IsVisible = bottomHasData;
                        cell.Borders.Left.IsVisible = leftHasData;
                        cell.Borders.Right.IsVisible = rightHasData;
                    }
                }
            }
        }

        private void SetCellBorders(Cell cell, bool isVisible, AColor color)
        {
            cell.Borders.Top.IsVisible = isVisible;
            cell.Borders.Bottom.IsVisible = isVisible;
            cell.Borders.Left.IsVisible = isVisible;
            cell.Borders.Right.IsVisible = isVisible;
            if (isVisible)
            {
                cell.Borders.Top.LineWeight = LineWeight.LineWeight025;
                cell.Borders.Bottom.LineWeight = LineWeight.LineWeight025;
                cell.Borders.Left.LineWeight = LineWeight.LineWeight025;
                cell.Borders.Right.LineWeight = LineWeight.LineWeight025;
                cell.Borders.Top.Color = color;
                cell.Borders.Bottom.Color = color;
                cell.Borders.Left.Color = color;
                cell.Borders.Right.Color = color;
            }
        }

        private void SwapDrillAttribute(BlockReference blockRef, string tag, string newValue, Transaction tr)
        {
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                if (attRef != null && attRef.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                {
                    attRef.TextString = newValue;
                }
            }
        }

        private void UpdateAttributeIfTagMatches(BlockReference blockRef, string tag, string newValue, Transaction tr, ref int updatedCount)
        {
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                if (attRef == null) continue;
                if (attRef.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.LogInfo($"Updating block attribute '{tag}' from '{attRef.TextString}' to '{newValue}' (Handle={blockRef.Handle}).");
                    attRef.TextString = newValue;
                    updatedCount++;
                }
            }
        }

        private AttributeReference GetAttributeReference(BlockReference blockRef, string tag, Transaction tr)
        {
            foreach (ObjectId attId in blockRef.AttributeCollection)
            {
                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                if (attRef != null && attRef.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                {
                    attRef.UpgradeOpen();
                    return attRef;
                }
            }
            return null;
        }

        private void RefreshUI()
        {
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                    for (int i = 0; i < DrillCount; i++)
                    {
                        string defaultName = GetDefaultDrillName(i);
                        string currentName = defaultName;
                        foreach (ObjectId entId in ms)
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent is BlockReference blockRef)
                            {
                                AttributeReference attRef = GetAttributeReference(blockRef, defaultName, tr);
                                if (attRef != null && !string.IsNullOrWhiteSpace(attRef.TextString))
                                {
                                    currentName = attRef.TextString;
                                    break;
                                }
                            }
                        }
                        drillTextBoxes[i].Text = currentName;
                        drillLabels[i].Text = currentName;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error refreshing UI: {ex.Message}");
                MessageBox.Show($"An error occurred while refreshing the UI: {ex.Message}",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetDefaultDrillName(int index)
        {
            return $"DRILL_{index + 1}";
        }

        private void SetDrill(int index, bool showMessage = true)
        {
            string defaultName = GetDefaultDrillName(index);
            string newDrillName = drillTextBoxes[index].Text.Trim();
            string oldDrillName = drillLabels[index].Text.Trim();
            Logger.LogInfo($"SetDrill => from '{oldDrillName}' to '{newDrillName}' (Tag='{defaultName}').");

            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            try
            {
                using (DocumentLock dlock = doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // unlock CG-NOTES layer to avoid eOnLockedLayer
                        var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                        const string targetLayer = "CG-NOTES";
                        if (lt.Has(targetLayer))
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(lt[targetLayer], OpenMode.ForWrite);
                            if (ltr.IsLocked) ltr.IsLocked = false;
                        }
                        string newValue = newDrillName;
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        int updatedBlocks = 0;
                        foreach (ObjectId btrId in bt)
                        {
                            BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                            if (btr == null || btr.IsErased)
                                continue;
                            foreach (ObjectId entId in btr)
                            {
                                Entity ent;
                                try
                                {
                                    ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                                }
                                catch
                                {
                                    continue;
                                }
                                if (ent == null || ent.IsErased)
                                    continue;
                                if (ent is BlockReference blockRef)
                                {
                                    foreach (ObjectId attId in blockRef.AttributeCollection)
                                    {
                                        AttributeReference attRef;
                                        try
                                        {
                                            attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                        if (attRef == null || attRef.IsErased)
                                            continue;
                                        if (attRef.Tag.Equals(defaultName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Logger.LogInfo($"SetDrill => Setting '{attRef.Tag}' from '{attRef.TextString}' to '{newValue}'.");
                                            attRef.TextString = newValue;
                                            updatedBlocks++;
                                        }
                                    }
                                }
                            }
                        }
                        tr.Commit();
                        if (updatedBlocks > 0)
                        {
                            Logger.LogInfo($"SetDrill => updated {updatedBlocks} attribute(s) for {defaultName}.");
                            if (showMessage)
                            {
                                MessageBox.Show($"Updated {updatedBlocks} attribute(s) for {defaultName}.", "Set Drill", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            Logger.LogWarning($"No DRILL_x attributes found for {defaultName} to update.");
                            if (showMessage)
                            {
                                MessageBox.Show($"No DRILL_x attributes found for {defaultName} to update.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
                UpdateAttributesWithMatchingValue(oldDrillName, newDrillName);
                drillLabels[index].Text = newDrillName;
                SaveToJson();
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error setting {defaultName}: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred while setting drill attributes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private string[,] ReadExcelData(string excelFilePath, string range)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var ws = package.Workbook.Worksheets[0];
                var addr = new OfficeOpenXml.ExcelAddress(range);
                int rows = addr.Rows;
                int cols = addr.Columns;
                string[,] data = new string[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string cellValue = ws.Cells[addr.Start.Row + r, addr.Start.Column + c].Text;
                        data[r, c] = cellValue;
                    }
                }
                return data;
            }
        }

        private void InsertAndFormatTable(Point3d insertionPoint, string[,] cellData, string tableStyleName)
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                int rows = cellData.GetLength(0);
                int cols = cellData.GetLength(1);

                double[] columnWidths;
                if (cols == 12)
                {
                    columnWidths = new double[] { 100.0, 100.0, 60.0, 60.0, 80.0, 80.0, 80.0, 80.0, 80.0, 80.0, 80.0, 80.0 };
                }
                else
                {
                    columnWidths = new double[cols];
                    for (int i = 0; i < cols; i++)
                    {
                        columnWidths[i] = 80.0;
                    }
                }

                Table table = new Table();
                table.TableStyle = GetTableStyleId(db, tableStyleName, tr);
                table.SetSize(rows, cols);
                table.Position = insertionPoint;
                // Explicitly set the table's layer to "CG-NOTES"
                table.Layer = "CG-NOTES";

                double defaultRowHeight = 25.0;
                double emptyCellRowHeight = 125.0;
                for (int r = 0; r < rows; r++)
                {
                    bool hasEmpty = false;
                    for (int c = 0; c < cols; c++)
                    {
                        string val = cellData[r, c];
                        table.Cells[r, c].TextString = val;
                        if (string.IsNullOrWhiteSpace(val))
                            hasEmpty = true;
                    }
                    table.Rows[r].Height = hasEmpty ? emptyCellRowHeight : defaultRowHeight;
                }

                for (int c = 0; c < cols; c++)
                {
                    table.Columns[c].Width = columnWidths[c];
                }

                btr.AppendEntity(table);
                tr.AddNewlyCreatedDBObject(table, true);

                table.GenerateLayout();
                UnmergeAllCells(table);
                table.GenerateLayout();
                RemoveAllCellBorders(table);
                AddBordersForDataCells(table);
                table.RecomputeTableBlock(true);

                tr.Commit();
            }
        }
        #endregion

        private void HighlightTextbox(TextBox textbox, bool highlight)
        {
            if (highlight)
            {
                textbox.BackColor = System.Drawing.Color.LightCoral; // Highlight with a red color
            }
            else
            {
                textbox.BackColor = System.Drawing.Color.LightGreen; ; // Default color
            }
        }
        private void CheckButton_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Step 1: Prompt the user to select the data-linked table
                    // Existing code to select and process the table
                    PromptEntityOptions peo = new PromptEntityOptions("\nSelect the data-linked table:");
                    peo.SetRejectMessage("\nOnly table entities are allowed.");
                    peo.AddAllowedClass(typeof(Table), true);

                    PromptEntityResult per = ed.GetEntity(peo);

                    if (per.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nTable selection cancelled.");
                        return;
                    }

                    // Get the table object
                    Table table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;

                    if (table == null)
                    {
                        ed.WriteMessage("\nSelected entity is not a table.");
                        return;
                    }

                    // Extract data from the table
                    List<string> tableValues = ExtractTableValues(table);

                    tr.Commit();

                    // Step 2: Prompt the user to select blocks with DRILLNAME attribute
                    List<BlockAttributeData> blockDataList = SelectAndExtractBlockData();

                    if (blockDataList == null || blockDataList.Count == 0)
                    {
                        MessageBox.Show("No blocks selected or no blocks with 'DRILLNAME' attribute found.", "Check Results", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Step 3: Perform comparisons and report discrepancies
                    CompareDrillNamesWithTableAndBlocks(tableValues, blockDataList);
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error during check operation: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred during the check operation: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetUtmsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Collect grid labels from layer "Z-DRILL-POINT" (change this string if needed)
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;
                var gridData = new List<(string Label, double Northing, double Easting)>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    foreach (ObjectId id in ms)
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        // Now check on the new layer "Z-DRILL-POINT"
                        if (ent != null && ent.Layer.Equals("Z-DRILL-POINT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (ent is DBText dbText)
                            {
                                string textValue = dbText.TextString.Trim();
                                if (IsGridLabel(textValue))
                                {
                                    gridData.Add((textValue, dbText.Position.Y, dbText.Position.X));
                                }
                            }
                            else if (ent is MText mText)
                            {
                                string textValue = mText.Contents.Trim();
                                if (IsGridLabel(textValue))
                                {
                                    gridData.Add((textValue, mText.Location.Y, mText.Location.X));
                                }
                            }
                        }
                    }
                    tr.Commit();
                }

                // Use natural sorting so that labels sort numerically (e.g. B2 comes before B10)
                gridData = gridData.OrderBy(x => x.Label, new NaturalStringComparer()).ToList();

                // Ensure the C:\CORDS folder exists and write the CSV file
                string cordsDir = @"C:\CORDS";
                Directory.CreateDirectory(cordsDir);
                string csvPath = Path.Combine(cordsDir, "CORDS.csv");
                using (StreamWriter sw = new StreamWriter(csvPath, false))
                {
                    sw.WriteLine("Label,Northing,Easting");
                    foreach (var pt in gridData)
                    {
                        sw.WriteLine($"{pt.Label},{pt.Northing},{pt.Easting}");
                    }
                }

                MessageBox.Show($"UTM CSV created successfully at:\n{csvPath}",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in GetUtmsButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error generating UTMs CSV: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void AddDrillPtsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Prompt user for a letter (e.g. "A")
                string letter = ShowInputDialog("Enter letter for drill points:", "Add Drill Pts", "A");
                if (string.IsNullOrWhiteSpace(letter))
                {
                    MessageBox.Show("Letter cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Prompt the user to select a polyline
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polyline:");
                peo.SetRejectMessage("\nOnly polylines are allowed.");
                peo.AddAllowedClass(typeof(Polyline), false);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nSelection cancelled.");
                    return;
                }

                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Ensure that the "Z-DRILL-POINT" layer exists.
                        EnsureLayer(db, "Z-DRILL-POINT");

                        Polyline pline = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
                        if (pline == null)
                        {
                            MessageBox.Show("Selected entity is not a polyline.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Limit to a maximum of 150 vertices
                        int count = Math.Min(pline.NumberOfVertices, 150);
                        for (int i = 0; i < count; i++)
                        {
                            Point2d pt2d = pline.GetPoint2dAt(i);
                            Point3d pt3d = new Point3d(pt2d.X, pt2d.Y, 0.0);
                            string label = $"{letter}{i + 1}";

                            // Create a DBText entity with specified properties
                            DBText dbt = new DBText();
                            dbt.Position = pt3d;
                            dbt.Height = 2.0;
                            dbt.TextString = label;
                            // Set the layer to "Z-DRILL-POINT"
                            dbt.Layer = "Z-DRILL-POINT";
                            dbt.ColorIndex = 7; // White

                            // Add the DBText to Model Space
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                            ms.AppendEntity(dbt);
                            tr.AddNewlyCreatedDBObject(dbt, true);
                        }
                        tr.Commit();
                    }
                }
                MessageBox.Show("Drill points added successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in AddDrillPtsButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Error: {ex.Message}", "Add Drill Pts", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void UpdateOffsetsButton_Click(object sender, EventArgs e)
        {
            try
            {
                Logger.LogInfo("Update Offsets: operation started.");

                // 1) Get active AutoCAD document and editor.
                Document doc = AcApplication.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No active AutoCAD document.", "Update Offsets",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                Editor ed = doc.Editor;
                Database db = doc.Database;

                // 2) Prompt the user to select a table.
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect the table to update offsets:");
                peo.SetRejectMessage("\n**Error:** Only table entities are allowed.");
                peo.AddAllowedClass(typeof(Table), exactMatch: true);
                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nUpdate Offsets canceled by user.\n");
                    return;
                }

                using (DocumentLock docLock = doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Open the selected table.
                    Table table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                    if (table == null)
                    {
                        MessageBox.Show("The selected entity is not a table.", "Update Offsets",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // We want to check columns C and D (0-based indices 2 and 3)
                    int[] columnsToCheck = { 2, 3 };
                    var dataCells = new List<(int row, int col, double originalVal, string direction, string originalText)>();

                    // Regex to match formats like "318.65N", "318.65 N", or "+318.65 S"
                    Regex offsetRegex = new Regex(@"^\s*([+-]?\d+(\.\d+)?)\s*([NnSsEeWw])\s*$");

                    // Loop through each row for columns C and D.
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        foreach (int colIndex in columnsToCheck)
                        {
                            string cellText = table.Cells[r, colIndex].TextString?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(cellText)) continue;

                            Match match = offsetRegex.Match(cellText);
                            if (match.Success)
                            {
                                string numericStr = match.Groups[1].Value;
                                string direction = match.Groups[3].Value.ToUpper();
                                if (double.TryParse(numericStr, NumberStyles.Any,
                                            CultureInfo.InvariantCulture, out double originalVal))
                                {
                                    dataCells.Add((r, colIndex, originalVal, direction, cellText));
                                }
                            }
                        }
                    }

                    if (dataCells.Count == 0)
                    {
                        MessageBox.Show("No numeric offset values found in columns C or D of the selected table.",
                                        "Update Offsets", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Logger.LogInfo("Update Offsets: No data found in columns C/D.");
                        return;
                    }

                    // 3) Gather text objects on layer "P-Drill-Offset" from model space.
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    var offsetTextIds = new List<ObjectId>();

                    foreach (ObjectId entId in ms)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        if (ent.Layer.Equals("P-Drill-Offset", StringComparison.OrdinalIgnoreCase)
                            && (ent is DBText || ent is MText))
                        {
                            offsetTextIds.Add(entId);
                        }
                    }

                    int textCount = offsetTextIds.Count;
                    if (textCount == 0)
                    {
                        MessageBox.Show("No offset text found on layer \"P-Drill-Offset\".",
                                        "Update Offsets", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Logger.LogInfo("No text objects on layer P-Drill-Offset.");
                        return;
                    }

                    // 4) If more text objects exist than table cells, select them all and abort.
                    int dataCellCount = dataCells.Count;
                    if (textCount > dataCellCount)
                    {
                        ed.SetImpliedSelection(offsetTextIds.ToArray());
                        ed.WriteMessage($"\n**Warning:** Found {textCount} offset text objects for only {dataCellCount} table entries. " +
                                        "All offset texts have been selected. Please review and try again.\n");
                        Logger.LogWarning($"Update Offsets aborted: {textCount} offset texts for {dataCellCount} table cells.");
                        return;
                    }

                    // 5) Parse numeric values from each text object.
                    List<(double value, ObjectId id)> offsetValues = new List<(double, ObjectId)>();
                    foreach (ObjectId tid in offsetTextIds)
                    {
                        Entity ent = tr.GetObject(tid, OpenMode.ForRead) as Entity;
                        if (ent is DBText dbText)
                        {
                            string txt = dbText.TextString.Trim();
                            Logger.LogInfo($"DBText => '{txt}'");
                            if (double.TryParse(txt, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                                offsetValues.Add((val, tid));
                        }
                        else if (ent is MText mText)
                        {
                            // Log the raw MText for debugging.
                            string raw = mText.Contents;
                            Logger.LogInfo($"Raw MText.Contents => '{raw}'");

                            // Use regex to replace formatting codes like {\C1;120.00} with just "120.00"
                            string plain = Regex.Replace(raw, @"\{\\[^\}]+;([^\}]+)\}", "$1");
                            // Remove any leftover formatting codes (if any)
                            plain = Regex.Replace(plain, @"\\[a-zA-Z]+\s*", " ");
                            plain = plain.Trim();
                            Logger.LogInfo($"Cleaned MText => '{plain}'");

                            // Split on newlines (if any)
                            string[] lines = plain.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            bool foundNumeric = false;
                            foreach (string line in lines)
                            {
                                string candidate = line.Trim();
                                if (double.TryParse(candidate, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                                {
                                    offsetValues.Add((val, tid));
                                    foundNumeric = true;
                                    break;
                                }
                                else
                                {
                                    // Optional: try to extract a numeric substring if the whole line isn't numeric.
                                    Match m = Regex.Match(candidate, @"([+-]?\d+(\.\d+)?)");
                                    if (m.Success && double.TryParse(m.Groups[1].Value, NumberStyles.Any,
                                                                     CultureInfo.InvariantCulture, out double extractedVal))
                                    {
                                        offsetValues.Add((extractedVal, tid));
                                        foundNumeric = true;
                                        break;
                                    }
                                }
                            }
                            if (!foundNumeric)
                            {
                                Logger.LogInfo("No numeric value found in MText => skipping.");
                            }
                        }
                    }

                    if (offsetValues.Count == 0)
                    {
                        MessageBox.Show("Text found on \"P-Drill-Offset\" layer, but none contained valid numeric values.",
                                        "Update Offsets", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Logger.LogInfo("No numeric values in offset text objects.");
                        return;
                    }

                    // 6) For each data cell, find the offset with the SMALLEST difference (within ±10.0).
                    double tolerance = 10.0;
                    int updatedCount = 0;
                    int noMatchCount = 0;
                    table.UpgradeOpen();  // Open table for modification.

                    foreach (var cellInfo in dataCells)
                    {
                        double bestDiff = double.MaxValue;
                        int bestIndex = -1;

                        for (int i = 0; i < offsetValues.Count; i++)
                        {
                            double candidateVal = offsetValues[i].value;
                            double diff = Math.Abs(candidateVal - cellInfo.originalVal);
                            if (diff < bestDiff)
                            {
                                bestDiff = diff;
                                bestIndex = i;
                            }
                        }

                        // If a matching offset is found within tolerance:
                        if (bestIndex != -1 && bestDiff <= tolerance)
                        {
                            double matchedVal = offsetValues[bestIndex].value;
                            // Remove so the same offset text isn't reused for a different cell.
                            offsetValues.RemoveAt(bestIndex);

                            // Format the new text with one decimal place + space + original direction
                            string newText = matchedVal.ToString("F1", CultureInfo.InvariantCulture)
                                             + " " + cellInfo.direction;
                            table.Cells[cellInfo.row, cellInfo.col].TextString = newText;
                            updatedCount++;
                        }
                        else
                        {
                            // No match found: revert the cell text and mark background red.
                            table.Cells[cellInfo.row, cellInfo.col].TextString = cellInfo.originalText;
                            table.Cells[cellInfo.row, cellInfo.col].BackgroundColor =
                                Autodesk.AutoCAD.Colors.Color.FromColor(System.Drawing.Color.Red);
                            noMatchCount++;
                        }
                    }

                    // Recompute layout to apply changes.
                    table.GenerateLayout();
                    tr.Commit();

                    ed.WriteMessage($"\nUpdate Offsets completed: {updatedCount} cells updated, {noMatchCount} cells with no match.\n");
                    Logger.LogInfo($"Update Offsets done => {updatedCount} updated, {noMatchCount} no-match cells.");
                }
            }
            catch (System.Exception ex)
            {
                Logger.LogError($"Error in UpdateOffsetsButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Update Offsets",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (string.IsNullOrEmpty(x))
                    return string.IsNullOrEmpty(y) ? 0 : -1;
                if (string.IsNullOrEmpty(y))
                    return 1;

                // Assume the first character is a letter and the rest is a number.
                char letterX = x[0];
                char letterY = y[0];
                int letterComparison = letterX.CompareTo(letterY);
                if (letterComparison != 0)
                    return letterComparison;

                // Extract numeric parts and compare as numbers.
                string numXStr = x.Substring(1);
                string numYStr = y.Substring(1);
                if (int.TryParse(numXStr, out int numX) && int.TryParse(numYStr, out int numY))
                {
                    return numX.CompareTo(numY);
                }
                return string.Compare(x, y, StringComparison.Ordinal);
            }
        }
        #region Logging

        // Logging is handled by the static Logger class (configured via NLog).
        // Calls to Logger.LogInfo, Logger.LogWarning, Logger.LogError are scattered through the code.

        #endregion
        private string GetReportFilePath()
        {
            Document acDoc = AcApplication.DocumentManager.MdiActiveDocument;
            if (acDoc == null || string.IsNullOrEmpty(acDoc.Name))
            {
                MessageBox.Show("No active AutoCAD document found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw new InvalidOperationException("No active AutoCAD document found.");
            }
            string dwgDirectory = Path.GetDirectoryName(acDoc.Name);
            string drawingName = Path.GetFileNameWithoutExtension(acDoc.Name);
            string reportFilePath = Path.Combine(dwgDirectory, $"{drawingName}_CheckReport.txt");
            return reportFilePath;
        }
        private void SaveToJson()
        {
            string jsonFilePath = GetJsonFilePath();
            try
            {
                // Collect current drill names and heading combo selection
                string[] drillNames = new string[DrillCount];
                for (int i = 0; i < DrillCount; i++)
                {
                    drillNames[i] = drillTextBoxes[i].Text.Trim();
                }
                var data = new
                {
                    DrillNames = drillNames,
                    ComboBoxValue = headingComboBox.SelectedItem.ToString()
                };
                File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(data, Formatting.Indented));
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Failed to save JSON: {ex.Message}",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void LoadFromJson()
        {
            string jsonFilePath = GetJsonFilePath();
            try
            {
                if (File.Exists(jsonFilePath))
                {
                    // Read JSON file contents
                    string jsonData = File.ReadAllText(jsonFilePath);
                    // Deserialize JSON data dynamically
                    dynamic data = JsonConvert.DeserializeObject<dynamic>(jsonData);

                    // Load drill names
                    string[] drillNames = data.DrillNames.ToObject<string[]>();
                    for (int i = 0; i < DrillCount; i++)
                    {
                        if (i < drillNames.Length && !string.IsNullOrWhiteSpace(drillNames[i]))
                        {
                            drillTextBoxes[i].Text = drillNames[i];
                            drillLabels[i].Text = drillNames[i];
                        }
                        else
                        {
                            string defaultName = $"DRILL_{i + 1}";
                            drillTextBoxes[i].Text = defaultName;
                            drillLabels[i].Text = defaultName;
                        }
                    }

                    // Restore Heading ComboBox selection
                    if (data.ComboBoxValue != null)
                    {
                        headingComboBox.SelectedItem = data.ComboBoxValue.ToString();
                    }
                    else
                    {
                        headingComboBox.SelectedItem = "OTHER";
                    }
                }
                else
                {
                    // If config file not found, initialize defaults and save one
                    InitializeDefaultDrills();
                    headingComboBox.SelectedItem = "OTHER";
                    SaveToJson();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error loading JSON data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                InitializeDefaultDrills();
                headingComboBox.SelectedItem = "OTHER";
            }
        }
        private Table SelectCoordinateTable()
        {
            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            // Prompt the user to select a table with coordinate data
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect the table with coordinate data:");
            peo.SetRejectMessage("\nOnly table entities are allowed.");
            peo.AddAllowedClass(typeof(Table), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                Logger.LogInfo("User cancelled table selection.");
                return null;
            }
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                Table table = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Table;
                tr.Commit();
                return table;
            }
        }
        private void UpdateAttributesWithMatchingValue(string oldValue, string newValue)
        {
            string oldValTrim = oldValue.Trim();
            string newValTrim = newValue.Trim();
            Logger.LogInfo($"Updating attributes from '{oldValTrim}' to '{newValTrim}' (exact match, ignore case).");

            Document doc = AcApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            int updatedCount = 0;

            using (DocumentLock dlock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // unlock CG-NOTES layer to avoid eOnLockedLayer
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    const string targetLayer = "CG-NOTES";
                    if (lt.Has(targetLayer))
                    {
                        var ltr = (LayerTableRecord)tr.GetObject(lt[targetLayer], OpenMode.ForWrite);
                        if (ltr.IsLocked) ltr.IsLocked = false;
                    }
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt == null)
                        return;
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null || btr.IsErased)
                            continue;
                        foreach (ObjectId entId in btr)
                        {
                            Entity ent = null;
                            try
                            {
                                ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                if ((Autodesk.AutoCAD.Runtime.ErrorStatus)ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.WasErased)
                                {
                                    Logger.LogWarning($"Skipped an erased entity (Handle={entId.Handle}).");
                                    continue;
                                }
                                throw;
                            }
                            if (ent == null || ent.IsErased)
                                continue;
                            if (ent is BlockReference blockRef)
                            {
                                foreach (ObjectId attId in blockRef.AttributeCollection)
                                {
                                    AttributeReference attRef;
                                    try
                                    {
                                        attRef = tr.GetObject(attId, OpenMode.ForWrite, false) as AttributeReference;
                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                    {
                                        if ((Autodesk.AutoCAD.Runtime.ErrorStatus)ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.WasErased)
                                        {
                                            Logger.LogWarning("Skipping an erased attribute reference.");
                                            continue;
                                        }
                                        throw;
                                    }
                                    if (attRef == null || attRef.IsErased)
                                        continue;
                                    if (attRef.TextString.Trim().Equals(oldValTrim, StringComparison.OrdinalIgnoreCase))
                                    {
                                        attRef.TextString = newValTrim;
                                        updatedCount++;
                                    }
                                }
                            }
                            else if (ent is DBText dbText)
                            {
                                if (dbText.TextString.Trim().Equals(oldValTrim, StringComparison.OrdinalIgnoreCase))
                                {
                                    dbText.TextString = newValTrim;
                                    updatedCount++;
                                }
                            }
                            else if (ent is MText mText)
                            {
                                if (mText.Contents.Trim().Equals(oldValTrim, StringComparison.OrdinalIgnoreCase))
                                {
                                    mText.Contents = newValTrim;
                                    updatedCount++;
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
            }

            if (updatedCount > 0)
            {
                Logger.LogInfo($"Successfully updated {updatedCount} occurrences of '{oldValTrim}' to '{newValTrim}'.");
                MessageBox.Show($"Replaced '{oldValTrim}' with '{newValTrim}' in {updatedCount} place(s).",
                                "Update Attributes", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                Logger.LogWarning($"No attributes found matching '{oldValTrim}'.");
                MessageBox.Show($"No attributes found matching '{oldValTrim}' to update.",
                                "Update Attributes", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


    }
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Enable visual styles and set text rendering
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            // Run the main form
            System.Windows.Forms.Application.Run(new FindReplaceForm());
        }
    }
    public class FindReplaceCommand
    {
        // The CommandMethod attribute makes this function callable from the AutoCAD command line.
        [CommandMethod("drillprops")]
        [CommandMethod("drillnames")]
        public static void ShowFindReplaceForm()
        {
            FindReplaceForm form = new FindReplaceForm();
            // Show the form modelessly so AutoCAD remains usable while it is open
            AcApplication.ShowModelessDialog(form);
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
    public class BlockData
    {
        public int WellId { get; set; }
        public string DrillName { get; set; }
    }
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
    public class WellsData
    {
        [JsonProperty("wells")]
        public List<DrillData> Wells { get; set; } = new List<DrillData>();
    }
}
