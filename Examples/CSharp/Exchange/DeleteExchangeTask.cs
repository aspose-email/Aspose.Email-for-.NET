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
    class DeleteExchangeTask
    {
        public static void Run()
        {
            // Create instance of ExchangeClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Get all tasks info collection from exchange
            ExchangeMessageInfoCollection tasks = client.ListMessages(client.MailboxInfo.TasksUri);

            // Parse all the tasks info in the list
            foreach (ExchangeMessageInfo info in tasks)
            {
                // Fetch task from exchange using current task info
                ExchangeTask task = client.FetchTask(info.UniqueUri);

                // Check if the current task fulfills the search criteria
                if (task.Subject.Equals("test"))
                {
                    //Delete task from exchange
                    client.DeleteTask(task.UniqueUri, DeleteTaskOptions.DeletePermanently);
                }
            }

            Console.WriteLine(Environment.NewLine + "Task deleted on Exchange Server successfully.");
        }
    }
}
