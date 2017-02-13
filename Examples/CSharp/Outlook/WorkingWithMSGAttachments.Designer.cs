namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    partial class WorkingWithMSGAttachments
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
            this.lstAttachments = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.txtFrom = new System.Windows.Forms.TextBox();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.txtTo = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lstAttachments
            // 
            this.lstAttachments.FormattingEnabled = true;
            this.lstAttachments.Location = new System.Drawing.Point(31, 210);
            this.lstAttachments.Name = "lstAttachments";
            this.lstAttachments.Size = new System.Drawing.Size(101, 121);
            this.lstAttachments.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(235, 210);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(91, 30);
            this.button1.TabIndex = 1;
            this.button1.Text = "Add Attachment";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(238, 295);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(90, 36);
            this.button2.TabIndex = 2;
            this.button2.Text = "Remove Attachment";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(241, 34);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(88, 28);
            this.button3.TabIndex = 3;
            this.button3.Text = "Save MSG File";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txtFrom
            // 
            this.txtFrom.Location = new System.Drawing.Point(71, 21);
            this.txtFrom.Name = "txtFrom";
            this.txtFrom.Size = new System.Drawing.Size(147, 20);
            this.txtFrom.TabIndex = 4;
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(71, 159);
            this.txtBody.Name = "txtBody";
            this.txtBody.Size = new System.Drawing.Size(147, 20);
            this.txtBody.TabIndex = 5;
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(71, 107);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(147, 20);
            this.txtSubject.TabIndex = 6;
            // 
            // txtTo
            // 
            this.txtTo.Location = new System.Drawing.Point(71, 64);
            this.txtTo.Name = "txtTo";
            this.txtTo.Size = new System.Drawing.Size(147, 20);
            this.txtTo.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "From";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(20, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "To";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 110);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Subject";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 166);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Body";
            // 
            // WorkingWithMSGAttachments
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(438, 359);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtTo);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.txtBody);
            this.Controls.Add(this.txtFrom);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lstAttachments);
            this.Name = "WorkingWithMSGAttachments";
            this.Text = "WorkingWithMSGAttachments";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lstAttachments;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox txtFrom;
        private System.Windows.Forms.TextBox txtBody;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.TextBox txtTo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}