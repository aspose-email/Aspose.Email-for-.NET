using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.CSharp.Exchange
{
    class CreateExchangeTask
    {
        public static void Run()
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

            Console.WriteLine(Environment.NewLine + "Task created on Exchange Server successfully.");
        }
    }
}
