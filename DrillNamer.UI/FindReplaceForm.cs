using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using DrillNamer.Core;

namespace DrillNamer.UI;

/// <summary>
/// Basic form to find and replace attribute values.
/// </summary>
public class FindReplaceForm : Form
{
    private readonly TextBox _oldValue = new TextBox();
    private readonly TextBox _newValue = new TextBox();
    private readonly Button _update = new Button();

    public FindReplaceForm()
    {
        _update.Text = "Replace";
        _update.Click += OnUpdateClick;
        var layout = new FlowLayoutPanel { Dock = DockStyle.Fill };
        layout.Controls.AddRange(new Control[] { _oldValue, _newValue, _update });
        Controls.Add(layout);
    }

    private void OnUpdateClick(object? sender, EventArgs e)
    {
        UpdateAttributesWithMatchingValue(_oldValue.Text, _newValue.Text);
    }

    /// <summary>
    /// Replace attributes matching <paramref name="oldValue"/> with <paramref name="newValue"/>.
    /// </summary>
    public void UpdateAttributesWithMatchingValue(string oldValue, string newValue)
    {
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        int updated = 0;

        using (DocumentLock dlock = doc.LockDocument())
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId btrId in bt)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                foreach (ObjectId entId in btr)
                {
                    if (tr.GetObject(entId, OpenMode.ForRead) is BlockReference br)
                    {
                        foreach (ObjectId attId in br.AttributeCollection)
                        {
                            if (tr.GetObject(attId, OpenMode.ForRead, false) is AttributeReference attRef)
                            {
                                if (attRef.TextString.Trim().Equals(oldValue.Trim(), StringComparison.OrdinalIgnoreCase))
                                {
                                    LayerState.WithUnlocked(attRef.LayerId, () =>
                                    {
                                        if (!attRef.IsWriteEnabled)
                                        {
                                            attRef.UpgradeOpen();
                                        }
                                        attRef.TextString = newValue.Trim();
                                    });
                                    updated++;
                                }
                            }
                        }
                    }
                }
            }
            tr.Commit();
        }

        MessageBox.Show($"Updated {updated} attribute(s).", "Update Attributes", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
