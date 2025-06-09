using Autodesk.AutoCAD.Runtime;
using DrillNamer.UI;

namespace DrillNamer.Commands;

/// <summary>
/// AutoCAD command entry points.
/// </summary>
public class DrillCommands
{
    /// <summary>
    /// Shows the find/replace form.
    /// </summary>
    [CommandMethod("DRILLNAMES", CommandFlags.NoHistory)]
    public void ShowFindReplaceForm()
    {
        var form = new FindReplaceForm();
        Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowModelessDialog(form);
    }
}
