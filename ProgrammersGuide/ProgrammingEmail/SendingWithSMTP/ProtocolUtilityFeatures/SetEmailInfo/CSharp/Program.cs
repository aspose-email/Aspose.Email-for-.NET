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

namespace SetEmailInfo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Create an instance MailMessage class
            MailMessage msg = new MailMessage();

            //Setting Date 
            msg.Date = DateTime.Now;

            //Setting Priority
            msg.Priority = MailPriority.High;

            //Setting Sensitivity
            msg.Sensitivity = MailSensitivity.Normal;

            //use MailMessage properties like specify sender, recipient and message
            msg.To = "asposetest123@gmail.com";
            msg.From = "asposetest123@gmail.com";
            msg.Subject = "Test Email";
            msg.Body = "Hello World!";


            //Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();

            try
            {
                //Client.Send will send this message
                client.Send(msg);
                //Show Message Sent�E only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f");
            client.SecurityMode = SmtpSslSecurityMode.Explicit;
            client.EnableSsl = true;

            return client;
        }
    }
}