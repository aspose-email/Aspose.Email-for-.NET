using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;
using System.Configuration;
using System.Data;

namespace Aspose.Email.Examples.CSharp.SMTP
{
    class ManagingEmailAttachments
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "outputAttachments.msg";
            string dstEmailRemoved = dataDir + "outputAttachmentRemoved.msg";

            //Create an instance of MailMessage class
            MailMessage message = new MailMessage();

            //From
            message.From = "sender@sender.com";

            //to whom
            message.To.Add("receiver@gmail.com");



            // 1.
            // Attaching an Attachment to an Email

            //Adding 1st attachment
            //Create an instance of Attachment class
            Attachment attachment;

            //Load an attachment
            attachment = new Attachment(dataDir + "1.txt");

            //Add attachment in instance of MailMessage class
            message.Attachments.Add(attachment);

            //Add 2nd Attachment
            message.AddAttachment(new Attachment(dataDir + "1.jpg"));

            //Add 3rd Attachment
            message.AddAttachment(new Attachment(dataDir + "1.doc"));

            //Add 4th Attachment
            message.AddAttachment(new Attachment(dataDir + "1.rar"));

            //Add 5th Attachment
            message.AddAttachment(new Attachment(dataDir + "1.pdf"));

            //Save message to disk
            message.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);



            // 2.
            // Removing an Attachment

            //Remove attachment from your MailMessage
            message.Attachments.Remove(attachment);

            //Save message to disk after removing a single attachment 
            message.Save(dstEmailRemoved, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);



            // 3.
            // Display remaining attachment names from message and save attachments

            //Create a loop to display the no. of attachments present in email message
            foreach (Attachment atchmnt in message.Attachments)
            {

                //Save your attachments here
                atchmnt.Save(dataDir + "attachment_" + atchmnt.Name);

                //display the the attachment file name
                Console.WriteLine(atchmnt.Name);
            }

            Console.WriteLine(Environment.NewLine + "Attachments removed successfully from " + dstEmailRemoved);
        }
    }
}
