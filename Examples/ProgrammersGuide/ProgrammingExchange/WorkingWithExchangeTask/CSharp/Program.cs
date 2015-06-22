//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WorkingWithExchnageTasks
{
    class Program
    {
        static void Main(string[] args)
        {
            //Initialize your license here

            //now call any of the methods below as per your requirements. Please note that you need a valid Exchange URI, 
            //UserName, and password (and domain if applicable) to try these examples.
        }
        static void CreateExchangeTask()
        {
            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Create Exchange task object
            ExchangeTask task = new ExchangeTask();

            // Set task subject
            task.Subject = "New-Test";

            // Set task status to In progress
            task.Status = ExchangeTaskStatus.InProgress;

            // Create task on exchange
            client.CreateTask(client.MailboxInfo.TasksUri, task);
        }

        static void UpdateExchangeTask()
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
        }

        static void DeleteExchangeTask()
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
        }

        static void SendExchangeTask()
        {
            // Create instance of ExchangeClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            MailMessageLoadOptions loadOptions = new MailMessageLoadOptions();
            loadOptions.MessageFormat = MessageFormat.Msg;
            loadOptions.FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments;

            // load task from .msg file
            MailMessage eml = MailMessage.Load("task.msg", loadOptions);
            eml.From = "firstname.lastname@domain.com";
            eml.To.Clear();
            eml.To.Add(new Aspose.Email.Mail.MailAddress("firstname.lastname@domain.com"));

            client.Send(eml);
        }

        static void SpecifyTimeZoneForTask()
        {
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            client.TimezoneId = "Central Europe Standard Time";
        }

        //Available for Aspose.Email for .NET 5.4.0 and onwards
        static void SaveExchangeTaskToDisc()
        {
            string dataDir = Path.GetFullPath("../../../Data/");

            ExchangeTask task = new ExchangeTask();
            task.Subject = "EMAILNET-34759 - " + Guid.NewGuid();
            task.Status = ExchangeTaskStatus.InProgress;
            task.StartDate = new DateTime(2028, 10, 6, 12, 30, 00);
            task.DueDate = task.StartDate.AddDays(3);
            task.Save(dataDir + "ExchangeTask.msg");
        }
    }
}
