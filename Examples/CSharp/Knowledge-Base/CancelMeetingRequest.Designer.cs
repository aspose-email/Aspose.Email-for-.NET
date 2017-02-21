namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    partial class CancelMeetingRequest
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
            this.dgMeetings = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgMeetings)).BeginInit();
            this.SuspendLayout();
            // 
            // dgMeetings
            // 
            this.dgMeetings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgMeetings.Location = new System.Drawing.Point(12, 26);
            this.dgMeetings.Name = "dgMeetings";
            this.dgMeetings.Size = new System.Drawing.Size(632, 117);
            this.dgMeetings.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(352, 173);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(225, 37);
            this.button1.TabIndex = 1;
            this.button1.Text = "Cancel Meeting Request";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // CancelMeetingRequest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(691, 261);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dgMeetings);
            this.Name = "CancelMeetingRequest";
            this.Text = "CancelMeetingRequest";
            ((System.ComponentModel.ISupportInitialize)(this.dgMeetings)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgMeetings;
        private System.Windows.Forms.Button button1;
    }
}