// Demonstrates deleting a task.
//
// Tasks are items like messages, so DeleteItem removes them by URI with the usual
// DeletionOptions. The example creates two tasks in a uniquely named tasks folder,
// deletes one and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteExchangeTask
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.TasksUri, folderName, ExchangeFolderType.Task);

                try
                {
                    var obsolete = client.CreateTask(folder.Uri, new ExchangeTask { Subject = "Obsolete task" });
                    client.CreateTask(folder.Uri, new ExchangeTask { Subject = "Keep this task" });

                    client.DeleteItem(obsolete, DeletionOptions.DeletePermanently);
                    Console.WriteLine("Deleted 'Obsolete task'.");

                    Console.WriteLine("Left in the folder:");
                    foreach (ExchangeTask task in client.ListTasks(folder.Uri))
                        Console.WriteLine($"  {task.Subject}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
