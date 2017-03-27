using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    public partial class WorkingWithMSGAttachments : Form
    {
        public WorkingWithMSGAttachments()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ExStart:CreateMessagesWithAttachments
            // Open a file open dialog to select the attachment
            OpenFileDialog fd = new OpenFileDialog();
            // If user selected the file and clicked OK, add the file name and path to the list
            if (fd.ShowDialog() == DialogResult.OK)
            {
                lstAttachments.Items.Add(fd.FileName);
            }
            // ExEnd:CreateMessagesWithAttachments
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // ExStart:RemoveAttachment 
            // Check if the user has selected any item in the attachments listbox
            if (lstAttachments.SelectedItems.Count > 0)
            {
                // Remove the selected attachment from the listbox
                lstAttachments.Items.RemoveAt(lstAttachments.SelectedIndex);
            }
            // ExEnd:RemoveAttachment
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string dataDir = RunExamples.GetDataDir_Outlook();
            
            // ExStart:AddingMSGAttachments
            // File name for output MSG file
            string strMsgFile;

            // Open a save file dialog for saving the file and Add filter for msg files
            SaveFileDialog fd = new SaveFileDialog();
            fd.Filter = "Outlook Message files (*.msg)|*.msg|All files (*.*)|*.*";

            // If user pressed OK, save the file name + path
            if (fd.ShowDialog() == DialogResult.OK)
            {
                strMsgFile = fd.FileName;
            }
            else
            {
                // If user did not selected the file, return
                return;
            }

            // Create an instance of MailMessage class
            MailMessage mailMsg = new MailMessage();
            // Set from, to, subject and body properties
            mailMsg.From = txtFrom.Text;
            mailMsg.To = txtTo.Text;
            mailMsg.Subject = txtSubject.Text;
            mailMsg.Body = txtBody.Text;

            // Add the attachments
            foreach (string strFileName in lstAttachments.Items)
            {
                mailMsg.Attachments.Add(new Attachment(strFileName));
            }

            // Create an instance of MapiMessage class and pass MailMessage as argument
            MapiMessage outlookMsg = MapiMessage.FromMailMessage(mailMsg);
            outlookMsg.Save(strMsgFile);
            // ExEnd:AddingMSGAttachments
        }
    }
}
