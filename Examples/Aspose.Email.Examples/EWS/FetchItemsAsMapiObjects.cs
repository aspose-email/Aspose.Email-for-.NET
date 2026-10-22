// Demonstrates fetching items as MAPI objects, with all their Outlook properties.
//
// FetchMapiMessages returns MapiMessage objects - the same model Aspose.Email uses for
// .msg files - so every MAPI property is available and the item can be saved as .msg as
// it is. FetchMapiTasks does the same for tasks, returning MapiTask objects with status,
// dates and progress.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchItemsAsMapiObjects
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var messageUris = client.ListMessages(client.MailboxInfo.InboxUri, 3)
                    .Select(info => info.UniqueUri)
                    .ToList();

                Console.WriteLine("Messages from the Inbox:");
                foreach (var message in client.FetchMapiMessages(messageUris))
                {
                    Console.WriteLine($"  {message.Subject}");
                    Console.WriteLine($"    class: {message.MessageClass}, attachments: {message.Attachments.Count}, " +
                                      $"properties: {message.Properties.Count}");
                }

                var taskUris = client.ListItems(client.MailboxInfo.TasksUri).Take(5).ToList();

                Console.WriteLine($"\nTasks ({taskUris.Count} shown):");
                foreach (var task in client.FetchMapiTasks(taskUris))
                {
                    Console.WriteLine($"  {task.Subject}: {task.Status}, {task.PercentComplete}% done, " +
                                      $"due {task.DueDate:d}");
                }
            }
        }
    }
}
