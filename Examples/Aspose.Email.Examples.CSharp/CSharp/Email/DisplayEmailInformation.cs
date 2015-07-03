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

namespace CSharp.Email
{
    class DisplayEmailInformation
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "test.eml";

            MailMessage message;

            //Create MailMessage instance by loading an Eml file
            message = MailMessage.Load(dataDir + "test.eml", MailMessageLoadOptions.DefaultEml);

            System.Console.Write("From:");

            //Gets the sender info
            System.Console.WriteLine(message.From);
            System.Console.Write("To:");

            //Gets the recipient info
            System.Console.WriteLine(message.To);
            System.Console.Write("Subject:");

            //Gets the subject
            System.Console.WriteLine(message.Subject);
            System.Console.WriteLine("HtmlBody:");

            //Gets the htmlbody 
            System.Console.WriteLine(message.HtmlBody);
            System.Console.WriteLine("TextBody");

            //Gets the textbody
            System.Console.WriteLine(message.Body);

            Console.WriteLine(Environment.NewLine + "Displayed email information from " + dstEmail);
        }
    }
}
