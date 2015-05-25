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

namespace MultipleRecipients
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Declare msg as MailMessage instance
            MailMessage message = new MailMessage();

            //use MailMessage properties like specify sender, recipient and message
            
            //Specify the recipients mail addresses
            message.To.Add("receiver1@receiver.com");
            message.To.Add("receiver2@receiver.com");
            message.To.Add("receiver3@receiver.com");
            message.To.Add("receiver4@receiver.com");

            message.CC.Add("CC1@receiver.com");
            message.CC.Add("CC2@receiver.com");

            message.Bcc.Add("Bcc1@receiver.com");
            message.Bcc.Add("Bcc2@receiver.com");

            message.From = "newcustomeronnet@gmail.com";
            message.Subject = "Test Email";
            message.Body = "Hello World!";


            //Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();          

            try
            {
                //Client will send this message
                client.Send(message);
                //Show only if message sent successfully
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
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}