// Demonstrates moving messages around a mailbox through Graph: create, update, copy,
// move, mark read and delete.
//
// UpdateSettings is the interesting part - by default an update rewrites the whole
// message, and SkipAttachments leaves the attachments on the server alone so a change
// to the subject does not re-upload them.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphMessages
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            var from = ClientBuilder.GraphMailboxId;
            const string subject = "Graph example message";

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                var workFolder = client.CreateFolder("Aspose.Email messages");

                // Create a message in a folder rather than sending it.
                var message = new MapiMessage(from, "recipient@example.com", subject,
                    "Created through Graph.", OutlookMessageFormat.Unicode);

                client.CreateMessage(workFolder.ItemId, message);
                Console.WriteLine($"Created a message in {workFolder.DisplayName}");

                var info = client.ListMessages(workFolder.ItemId, null)
                    .Cast<MessageInfo>()
                    .FirstOrDefault(m => m.Subject == subject);

                if (info == null)
                {
                    Console.WriteLine("The message was not found after creating it.");
                    return;
                }

                Console.WriteLine($"Item id: {info.ItemId}");

                // Fetch, change, and push the change back.
                var fetched = client.FetchMessage(info.ItemId);
                fetched.Subject = subject + " (updated)";

                // Without SkipAttachments the attachments would be uploaded again.
                var updated = client.UpdateMessage(fetched, new UpdateSettings { SkipAttachments = true });
                Console.WriteLine($"Updated the subject to: {updated.Subject}");

                client.SetRead(info.ItemId);
                Console.WriteLine("Marked it as read.");

                // Copy and move take the destination folder first.
                var inbox = client.GetFolder(KnownFolders.Inbox);
                client.CopyMessage(inbox.ItemId, info.ItemId);
                Console.WriteLine("Copied it to the Inbox.");

                client.MoveMessage(inbox.ItemId, info.ItemId);
                Console.WriteLine("Moved the original to the Inbox as well.");

                // Tidy up.
                client.Delete(workFolder.ItemId);
                Console.WriteLine($"Deleted {workFolder.DisplayName}");
            }
        }
    }
}
