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
