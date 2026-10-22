// Demonstrates creating a task.
//
// ExchangeTask holds the task's subject, status, dates, priority and body; CreateTask
// stores it in a tasks folder and returns its URI. The example creates the task in a
// uniquely named tasks folder, lists it and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateExchangeTask
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.TasksUri, folderName, ExchangeFolderType.Task);

                try
                {
                    var task = new ExchangeTask
                    {
                        Subject = "Prepare the quarterly report",
                        Status = ExchangeTaskStatus.InProgress,
                        StartDate = DateTime.Today,
                        DueDate = DateTime.Today.AddDays(7),
                        Priority = MailPriority.High
                    };

                    var uri = client.CreateTask(folder.Uri, task);
                    Console.WriteLine($"Created task: {uri}");

                    foreach (ExchangeTask stored in client.ListTasks(folder.Uri))
                        {
                            Console.WriteLine($"  {stored.Subject}: {stored.Status}, " +
                                              $"due {stored.DueDate:d}, {stored.Priority}");
                        }
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
