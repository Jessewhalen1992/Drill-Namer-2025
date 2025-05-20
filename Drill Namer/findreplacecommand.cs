using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application; // Alias for AutoCAD Application
using System.Windows.Forms; // Use this for Windows Forms

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
