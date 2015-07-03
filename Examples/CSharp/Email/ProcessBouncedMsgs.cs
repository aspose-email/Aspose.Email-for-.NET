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
using System.Data;
using Aspose.Email.Mail.Bounce;

namespace CSharp.Email
{
    class ProcessBouncedMsgs
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "test.eml";

            string fileName = dstEmail;
            MailMessage mail = MailMessage.Load(fileName);
            BounceResult result = mail.CheckBounced();
            Console.WriteLine(fileName);
            Console.WriteLine("IsBounced : " + result.IsBounced);
            Console.WriteLine("Action : " + result.Action);
            Console.WriteLine("Recipient : " + result.Recipient);
            
            Console.WriteLine(Environment.NewLine + "Bounce information displayed successfully.");
        }
    }
}
