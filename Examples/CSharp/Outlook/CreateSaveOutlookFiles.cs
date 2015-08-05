//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Outlook.Pst;
using Aspose.Email.Outlook;
using Aspose.Email.Mail;

namespace CSharp.Outlook
{
    class CreateSaveOutlookFiles
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "message.msg";

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
            outlookMsg.Save(dst);

            Console.WriteLine(Environment.NewLine + "MSG saved successfully at " + dst);
        }
    }
}
