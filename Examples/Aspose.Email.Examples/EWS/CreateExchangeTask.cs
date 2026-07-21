using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class CreateExchangeTask
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {

                // Create Exchange task object
                var task = new ExchangeTask
                {
                    // Set task subject
                    Subject = "New-Task",

                    // Set task status to In progress
                    Status = ExchangeTaskStatus.InProgress
                };

                // Create task on exchange
                client.CreateTask(client.MailboxInfo.TasksUri, task);
                Console.WriteLine("Task created on Exchange Server successfully.");
            }
        }
    }
}
