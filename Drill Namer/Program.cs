using System;
using System.Windows.Forms;

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
