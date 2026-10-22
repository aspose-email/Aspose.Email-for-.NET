// Demonstrates cleaning up tasks: deleting every completed task in a folder.
//
// A TaskStatus query finds the completed tasks, and DeleteItems removes them in one
// request - MoveToDeletedItems keeps them recoverable. The example works in a uniquely
// named tasks folder and deletes it at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteTaskOnExchange
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.TasksUri, folderName, ExchangeFolderType.Task);

                try
                {
                    client.CreateTask(folder.Uri,
                        new ExchangeTask { Subject = "Done 1", Status = ExchangeTaskStatus.Completed });
                    client.CreateTask(folder.Uri,
                        new ExchangeTask { Subject = "Done 2", Status = ExchangeTaskStatus.Completed });
                    client.CreateTask(folder.Uri,
                        new ExchangeTask { Subject = "Still open", Status = ExchangeTaskStatus.InProgress });

                    var builder = new ExchangeQueryBuilder();
                    builder.TaskStatus.Equals(ExchangeTaskStatus.Completed);

                    var completed = client.ListTasks(folder.Uri, builder.GetQuery()).Cast<ExchangeTask>().ToList();
                    client.DeleteItems(completed.Select(task => task.UniqueUri), DeletionOptions.MoveToDeletedItems);
                    Console.WriteLine($"Moved {completed.Count} completed task(s) to Deleted Items.");

                    Console.WriteLine("Left in the folder:");
                    foreach (ExchangeTask task in client.ListTasks(folder.Uri))
                        Console.WriteLine($"  {task.Subject}: {task.Status}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
