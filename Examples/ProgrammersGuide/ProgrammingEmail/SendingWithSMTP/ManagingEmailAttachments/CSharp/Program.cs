//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Mail;
using System;

namespace ManagingEmailAttachments
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

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
            message.Save(dataDir + "outputAttachments.msg", MessageFormat.Msg);



            // 2.
            // Removing an Attachment

            //Remove attachment from your MailMessage
            message.Attachments.Remove(attachment);

            //Save message to disk after removing a single attachment 
            message.Save(dataDir + "outputAttachmentRemoved.msg", MessageFormat.Msg);



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
        }
    }
}