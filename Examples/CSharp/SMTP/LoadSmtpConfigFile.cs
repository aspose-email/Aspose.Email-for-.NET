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
using System.Configuration;

namespace CSharp.SMTP
{
    class LoadSmtpConfigFile
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "EmbeddedImage.msg";

            //Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            //use MailMessage properties like specify sender, recipient and message
            msg.To = "asposetest123@gmail.com";
            msg.From = "aspose2@gmail.com";
            msg.Subject = "Test Email";
            msg.Body = "Hello World!";

            //Create an instance of SmtpClient class and load SMTP Authentication settings from Config file
            SmtpClient client = new SmtpClient(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None));
            client.SecurityOptions = SecurityOptions.Auto;

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

            Console.WriteLine(Environment.NewLine + "Message sent after loading configuration from config file.");
        }
    }
}
