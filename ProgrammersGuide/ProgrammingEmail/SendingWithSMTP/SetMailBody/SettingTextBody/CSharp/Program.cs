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

namespace SettingTextBody
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
            //use MailMessage properties like specify sender, recipient and message
            msg.From = "newcustomeronnet@gmail.com";
            msg.To = "newcustomeronnet2@gmail.com";
            msg.Subject = "Test subject";
            msg.TextBody = "This is text body";


            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f");

            client.SecurityMode = SmtpSslSecurityMode.Explicit;

            client.EnableSsl = true;

            try
            {
                //Client will send this message
                client.Send(msg);
                //Show only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
            
        }
    }
}