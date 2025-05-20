using System;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

namespace Drill_Namer
{
    public partial class FindReplaceForm : Form
    {
        private TextBox[] drillTextBoxes;
        private Label[] drillLabels;
        private const string DataFilePath = "drill_data.json";

        public FindReplaceForm()
        {
            InitializeComponent();
            InitializeDynamicControls();
            LoadData(); // Load JSON data on form load
        }

        private void InitializeDynamicControls()
        {
            this.ClientSize = new System.Drawing.Size(700, 700);

            drillTextBoxes = new TextBox[12];
            drillLabels = new Label[12];
            Button[] setButtons = new Button[12];
            Button[] updateButtons = new Button[12];

            for (int i = 0; i < 12; i++)
            {
                drillLabels[i] = new Label();
                drillLabels[i].Text = $"DRILL_{i + 1}";
                drillLabels[i].Location = new System.Drawing.Point(10, 20 + i * 35);
                drillLabels[i].Size = new System.Drawing.Size(150, 25);

                drillTextBoxes[i] = new TextBox();
                drillTextBoxes[i].Location = new System.Drawing.Point(180, 20 + i * 35);
                drillTextBoxes[i].Size = new System.Drawing.Size(200, 25);

                setButtons[i] = new Button();
                setButtons[i].Text = "SET";
                setButtons[i].Location = new System.Drawing.Point(400, 20 + i * 35);
                int currentIndex = i;
                setButtons[i].Click += (sender, e) => SetButton_Click(currentIndex);

                updateButtons[i] = new Button();
                updateButtons[i].Text = "CREATE";
                updateButtons[i].Location = new System.Drawing.Point(460, 20 + i * 35);
                updateButtons[i].Click += (sender, e) => CreateButton_Click(currentIndex);

                this.Controls.Add(drillLabels[i]);
                this.Controls.Add(drillTextBoxes[i]);
                this.Controls.Add(setButtons[i]);
                this.Controls.Add(updateButtons[i]);
            }

            Button setAllButton = new Button();
            setAllButton.Text = "SET ALL";
            setAllButton.Location = new System.Drawing.Point(180, 480);
            setAllButton.Click += (sender, e) => SetAllButton_Click();

            Button resetAllButton = new Button();
            resetAllButton.Text = "RESET ALL";
            resetAllButton.Location = new System.Drawing.Point(260, 480);
            resetAllButton.Click += (sender, e) => ResetAllButton_Click();

            Button loadJsonButton = new Button();
            loadJsonButton.Text = "LOAD JSON";
            loadJsonButton.Location = new System.Drawing.Point(340, 480);
            loadJsonButton.Click += (sender, e) => LoadData();

            this.Controls.Add(setAllButton);
            this.Controls.Add(resetAllButton);
            this.Controls.Add(loadJsonButton);
        }

        private void SetButton_Click(int index)
        {
            if (index >= 0 && index < drillTextBoxes.Length)
            {
                string userInput = drillTextBoxes[index].Text;
                ReplaceTextInAutoCAD($"DRILL_{index + 1}", userInput);
                drillLabels[index].Text = userInput;
                SaveData();
            }
        }

        private void ResetAllButton_Click()
        {
            DialogResult result = MessageBox.Show("Are you sure you want to reset all drill names to defaults?", "Confirm Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                for (int i = 0; i < 12; i++)
                {
                    drillLabels[i].Text = $"DRILL_{i + 1}";
                    drillTextBoxes[i].Text = "";
                }
                SaveData();
            }
        }

        private void SetAllButton_Click()
        {
            for (int i = 0; i < 12; i++)
            {
                SetButton_Click(i);
            }
            MessageBox.Show("All fields updated.");
        }

        private void ReplaceTextInAutoCAD(string findText, string replaceText)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            int replacementCount = 0;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entId in btr)
                    {
                        try
                        {
                            Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent == null) continue;

                            if (ent is DBText dbText && dbText.TextString.Contains(findText))
                            {
                                if (!dbText.IsWriteEnabled) dbText.UpgradeOpen();
                                dbText.TextString = dbText.TextString.Replace(findText, replaceText);
                                replacementCount++;
                            }
                            else if (ent is MText mText && mText.Contents.Contains(findText))
                            {
                                if (!mText.IsWriteEnabled) mText.UpgradeOpen();
                                mText.Contents = mText.Contents.Replace(findText, replaceText);
                                replacementCount++;
                            }
                            else if (ent is BlockReference blockRef)
                            {
                                foreach (ObjectId attId in blockRef.AttributeCollection)
                                {
                                    AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                    if (attRef != null && attRef.TextString.Contains(findText))
                                    {
                                        if (!attRef.IsWriteEnabled) attRef.UpgradeOpen();
                                        attRef.TextString = attRef.TextString.Replace(findText, replaceText);
                                        replacementCount++;
                                    }
                                }
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            ed.WriteMessage($"\nError processing entity {entId}: {ex.Message}");
                        }
                    }
                }
                tr.Commit();
            }

            ed.WriteMessage($"\nReplaced '{findText}' with '{replaceText}' in {replacementCount} entities.");
        }

        private void SaveData()
        {
            string[] data = new string[12];
            for (int i = 0; i < 12; i++)
            {
                if (string.IsNullOrEmpty(drillLabels[i].Text))
                {
                    data[i] = $"DRILL_{i + 1}"; // Default to DRILL_# if no user input
                }
                else
                {
                    data[i] = drillLabels[i].Text;
                }
            }

            string filePath = Path.ChangeExtension(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name, ".json");
            File.WriteAllText(filePath, JsonConvert.SerializeObject(data));
        }

        private void LoadData()
        {
            string filePath = Path.ChangeExtension(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name, ".json");
            if (File.Exists(filePath))
            {
                string[] data = JsonConvert.DeserializeObject<string[]>(File.ReadAllText(filePath));
                for (int i = 0; i < 12; i++)
                {
                    drillLabels[i].Text = string.IsNullOrEmpty(data[i]) ? $"DRILL_{i + 1}" : data[i]; // Set default if data is empty
                }
            }
            else
            {
                MessageBox.Show("No JSON file found for this drawing. Please create a new one.");
            }
        }

        private void CreateButton_Click(int index)
        {
            MessageBox.Show($"CREATE clicked for DRILL_{index + 1}");
            // Add your CREATE layout logic here
        }
    }
}
