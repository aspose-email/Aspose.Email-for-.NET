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
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using System.Net;

namespace CSharp.Exchange
{
    class SendExchangeTask
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Exchange();
            string dstEmail = dataDir + "Message.msg";
            
            // Create instance of ExchangeClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            MailMessageLoadOptions loadOptions = new MailMessageLoadOptions();
            loadOptions.MessageFormat = MessageFormat.Msg;
            loadOptions.FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments;

            // load task from .msg file
            MailMessage eml = MailMessage.Load(dstEmail, loadOptions);
            eml.From = "firstname.lastname@domain.com";
            eml.To.Clear();
            eml.To.Add(new Aspose.Email.Mail.MailAddress("firstname.lastname@domain.com"));

            client.Send(eml);

            Console.WriteLine(Environment.NewLine + "Task sent on Exchange Server successfully.");
        }
    }
}
