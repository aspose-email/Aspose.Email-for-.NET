using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.EWS
{
    class ProcessExchangeTasksUsingIEWSClient
    {
        public static void Run()
        {
            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Create Exchange task object
            ExchangeTask task = new ExchangeTask();

            // Set task subject and status to In progress
            task.Subject = "New-Test";
            task.Status = ExchangeTaskStatus.InProgress;
            client.CreateTask(client.MailboxInfo.TasksUri, task);
        }
    }
}