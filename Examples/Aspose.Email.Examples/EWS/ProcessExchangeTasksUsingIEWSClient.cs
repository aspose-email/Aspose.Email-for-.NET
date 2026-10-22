// Demonstrates a typical task-processing loop: find the tasks in a given state and move
// them on.
//
// ListTasks with a TaskStatus query returns the matching tasks; change each one and
// write it back with UpdateTask. Here every task that is in progress is completed. The
// example works in a uniquely named tasks folder and deletes it at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ProcessExchangeTasksUsingIEWSClient
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
                        new ExchangeTask { Subject = "Write tests", Status = ExchangeTaskStatus.InProgress });
                    client.CreateTask(folder.Uri,
                        new ExchangeTask { Subject = "Review code", Status = ExchangeTaskStatus.InProgress });
                    client.CreateTask(folder.Uri,
                        new ExchangeTask { Subject = "Plan release", Status = ExchangeTaskStatus.NotStarted });

                    var builder = new ExchangeQueryBuilder();
                    builder.TaskStatus.Equals(ExchangeTaskStatus.InProgress);

                    foreach (ExchangeTask task in client.ListTasks(folder.Uri, builder.GetQuery()))
                    {
                        task.Status = ExchangeTaskStatus.Completed;
                        client.UpdateTask(task);
                        Console.WriteLine($"Completed '{task.Subject}'.");
                    }

                    Console.WriteLine("\nAll tasks now:");
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
