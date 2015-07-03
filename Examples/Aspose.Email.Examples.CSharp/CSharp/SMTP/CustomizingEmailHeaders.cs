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

namespace CSharp.SMTP
{
    class CustomizingEmailHeaders
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "MsgHeaders.msg";

            //Create an instance MailMessage class
            MailMessage msg = new MailMessage();

            //Specify ReplyTo
            msg.ReplyToList.Add("reply@reply.com");

            //From field
            msg.From = "sender@sender.com";

            //To field
            msg.To.Add("receiver1@receiver.com");

            //Adding Cc and Bcc Addresses
            msg.CC.Add("receiver2@receiver.com");
            msg.Bcc.Add("receiver3@receiver.com");

            //Message subject
            msg.Subject = "test mail";

            //Specify Date
            msg.Date = new System.DateTime(2006, 3, 6);

            //Specify XMailer
            msg.XMailer = "Aspose.Email";

            //Specify Secret Header
            msg.Headers.Add("secret-header", "mystery");

            //Save message to disc
            msg.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);

            Console.WriteLine(Environment.NewLine + "Message saved with customized headers successfully at " + dstEmail);
        }
    }
}
