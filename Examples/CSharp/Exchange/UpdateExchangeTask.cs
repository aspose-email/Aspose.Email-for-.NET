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
    class UpdateExchangeTask
    {
        public static void Run()
        {
            // Create and initialize credentials
            var credentials = new NetworkCredential("username", "12345");

            // Create instance of ExchangeClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Get all tasks info collection from exchange
            ExchangeMessageInfoCollection tasks = client.ListMessages(client.MailboxInfo.TasksUri);

            // Parse all the tasks info in the list
            foreach (ExchangeMessageInfo info in tasks)
            {
                // Fetch task from exchange using current task info
                ExchangeTask task = client.FetchTask(info.UniqueUri);

                // Update the task status to NotStarted
                task.Status = ExchangeTaskStatus.NotStarted;

                // Set the task due date
                task.DueDate = new DateTime(2013, 2, 26);

                // Set task priority
                task.Priority = MailPriority.Low;

                // Update task on exchange
                client.UpdateTask(task);
            }

            Console.WriteLine(Environment.NewLine + "Task updated on Exchange Server successfully.");
        }
    }
}
