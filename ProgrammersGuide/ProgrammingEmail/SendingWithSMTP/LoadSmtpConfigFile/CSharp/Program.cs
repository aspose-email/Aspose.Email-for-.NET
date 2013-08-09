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
using System.Configuration;

namespace LoadSmtpConfigFile
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            //use MailMessage properties like specify sender, recipient and message
            msg.To = "asposetest123@gmail.com";
            msg.From = "aspose2@gmail.com";
            msg.Subject = "Test Email";
            msg.Body = "Hello World!";

            //Create an instance of SmtpClient class and load SMTP Authentication settings from Config file
            SmtpClient client = new SmtpClient(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None));
            client.EnableSsl = true;

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
        }
    }
}