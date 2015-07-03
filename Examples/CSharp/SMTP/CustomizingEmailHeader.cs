//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace CSharp.SMTP
{
    class CustomizingEmailHeader
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test.eml";

            //Create an instance MailMessage class
            MailMessage msg = new MailMessage();

            //Specify ReplyTo
            msg.ReplyToList.Add("reply@reply.com");

            //From field
            msg.From = "sender@sender.com";

            //To field
            msg.To.Add("receiver1@receiver.com");

            //Adding CC and BCC Addresses
            msg.CC.Add("receiver2@receiver.com");
            msg.Bcc.Add("receiver3@receiver.com");

            //Message subject
            msg.Subject = "test mail";

            //Specify Date
            msg.Date = new System.DateTime(2006, 3, 6);

            //Specify XMailer
            msg.XMailer = "Aspose.Email";

            msg.Headers.Add_("secret-header", "mystery");

            //Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();

            try
            {
                //Client.Send will send this message
                client.Send(msg);
                //Message sent successfully
                System.Console.WriteLine("Message sent");
            }

            catch (System.Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Email sent with customized headers.");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}
