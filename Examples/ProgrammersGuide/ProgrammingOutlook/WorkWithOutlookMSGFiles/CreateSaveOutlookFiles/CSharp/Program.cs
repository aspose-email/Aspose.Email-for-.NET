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
using Aspose.Email.Outlook;

namespace CreateSaveOutlookFiles
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            Directory.CreateDirectory(dataDir);

            // Create an instance of MailMessage class
            MailMessage mailMsg = new MailMessage();
            
            // Set FROM field of the message
            mailMsg.From = "from@domain.com";

            // Set TO field of the message
            mailMsg.To.Add("to@domain.com");
            
            // Set SUBJECT of the message
            mailMsg.Subject = "creating an outlook message file";

            // Set BODY of the message
            mailMsg.Body = "This message is created by Aspose.Email";

            // Create an instance of MapiMessage class and pass MailMessage as argument
            MapiMessage outlookMsg = MapiMessage.FromMailMessage(mailMsg);
            
            // Save the message (msg) file
            outlookMsg.Save(dataDir + "message.msg");

            // Display Status.
            System.Console.WriteLine("MSG file created and saved successfully.");
        }
    }
}