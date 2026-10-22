// Demonstrates changing an existing task.
//
// FetchTask reads the task by URI; change its status, dates or priority and pass it to
// UpdateTask. The example creates a task in a uniquely named tasks folder, updates it
// and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class UpdateExchangeTask
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.TasksUri, folderName, ExchangeFolderType.Task);

                try
                {
                    var uri = client.CreateTask(folder.Uri, new ExchangeTask
                    {
                        Subject = "Renew the domain",
                        Status = ExchangeTaskStatus.NotStarted,
                        DueDate = DateTime.Today.AddDays(30)
                    });

                    var task = client.FetchTask(uri);
                    Console.WriteLine($"Before: {task.Status}, due {task.DueDate:d}, {task.Priority}");

                    task.Status = ExchangeTaskStatus.InProgress;
                    task.DueDate = DateTime.Today.AddDays(7);
                    task.Priority = MailPriority.High;
                    client.UpdateTask(task);

                    var updated = client.FetchTask(uri);
                    Console.WriteLine($"After:  {updated.Status}, due {updated.DueDate:d}, {updated.Priority}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
