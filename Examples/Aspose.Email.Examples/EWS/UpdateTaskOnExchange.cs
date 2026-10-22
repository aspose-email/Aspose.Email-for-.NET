// Demonstrates updating a task through its MAPI representation.
//
// The UpdateTask overload that takes a MapiTask writes Outlook task properties - such as
// percent complete and status - that the ExchangeTask model does not cover, and returns
// the task's URI. The example creates a task in a uniquely named tasks folder, marks it
// half done, and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class UpdateTaskOnExchange
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.TasksUri, folderName, ExchangeFolderType.Task);

                try
                {
                    var uri = client.CreateTask(folder.Uri, new ExchangeTask { Subject = "Migrate the wiki" });

                    var task = (MapiTask)client.FetchItem(uri).ToMapiMessageItem();
                    Console.WriteLine($"Before: {task.Status}, {task.PercentComplete}% done");

                    task.Status = MapiTaskStatus.InProgress;
                    task.PercentComplete = 50;
                    client.UpdateTask(uri, task);

                    var updated = (MapiTask)client.FetchItem(uri).ToMapiMessageItem();
                    Console.WriteLine($"After:  {updated.Status}, {updated.PercentComplete}% done");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}
