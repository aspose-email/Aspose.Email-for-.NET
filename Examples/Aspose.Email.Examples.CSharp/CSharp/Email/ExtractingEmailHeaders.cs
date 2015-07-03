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
    class ExtractingEmailHeaders
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "test.eml";

            MailMessage message;

            //Create MailMessage instance by loading an EML file
            message = MailMessage.Load(dataDir + "email-headers.eml", MailMessageLoadOptions.DefaultEml);

            Console.WriteLine("\n\nheaders:\n\n");

            //print out all the headers
            int index = 0;
            foreach (String header in message.Headers)
            {
                Console.Write(header + " - ");
                Console.WriteLine(message.Headers.Get(index++));//.GetValues(header).Length.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Displayed email headers from " + dstEmail);
        }
    }
}
