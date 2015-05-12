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

namespace SSLEnabledSMTPServer
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            Aspose.Email.Mail.SmtpClient client = new Aspose.Email.Mail.SmtpClient("smtp.gmail.com");

            // Set username
            client.Username = "asposetest123@gmail.com";
            
            // Set password
            client.Password = "F123456f";
            
            // Set the port to 587. This is the SSL port of SMTP server
            client.Port = 587;
            
            // Set the security mode to explicit
            client.SecurityMode = SmtpSslSecurityMode.Explicit;
            
            // Enable SSL
            client.EnableSsl = true;

            //Declare msg as MailMessage instance
            MailMessage msg = new MailMessage();

            //use MailMessage properties like specify sender, recipient and message
            msg.To = "newcustomeronnet@gmail.com";
            msg.From = "newcustomeronnet@gmail.com";
            msg.Subject = "Test Email";
            msg.Body = "Hello World!";

            try
            {
                //Client will send this message
                client.Send(msg);
                //Show only if message sent successfully
                System.Console.WriteLine("Message sent");
            }

            catch (System.Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
           
        }
    }
}