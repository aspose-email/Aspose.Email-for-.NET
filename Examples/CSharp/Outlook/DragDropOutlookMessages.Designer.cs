namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    partial class DragDropOutlookMessages
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.myPanel1 = new Aspose.Email.Examples.CSharp.Email.Outlook.MyPanel();
            this.SuspendLayout();
            // 
            // myPanel1
            // 
            // ExStart:AllowDrop
            this.myPanel1.AllowDrop = true;
            // ExEnd:AllowDrop
            this.myPanel1.Location = new System.Drawing.Point(79, 49);
            this.myPanel1.Name = "myPanel1";
            this.myPanel1.Size = new System.Drawing.Size(200, 100);
            this.myPanel1.TabIndex = 0;
            this.myPanel1.DragDrop += new System.Windows.Forms.DragEventHandler(this.myPanel1_DragDrop);
            // 
            // DragDropOutlookMessages
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.myPanel1);
            this.Name = "DragDropOutlookMessages";
            this.Text = "DragDropOutlookMessages";
            this.ResumeLayout(false);

        }

        #endregion

        private Aspose.Email.Examples.CSharp.Email.Outlook.MyPanel myPanel1;
    }
}