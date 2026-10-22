// Demonstrates getting notified about changes in a mailbox folder.
//
// Subscribe to one of the *FolderServerNotifications events (here the Inbox) and choose
// which changes to hear about with the matching *FolderEventFilter - new mail, created,
// deleted, moved items and so on. NotificationsCheckInterval and NotificationTimeout
// tune how the client waits for them. To have something to report, the example appends
// a test message to the Inbox and deletes it again afterwards.

using System;
using System.Threading;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class MonitorExchangeFolderNotifications
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            using (var notified = new ManualResetEventSlim(false))
            {
                client.InboxFolderEventFilter = EventTypes.NewMail | EventTypes.Created | EventTypes.Deleted;
                client.NotificationsCheckInterval = 5000;  // milliseconds
                client.NotificationTimeout = 60000;        // milliseconds

                client.InboxFolderServerNotifications += (sender, e) =>
                {
                    Console.WriteLine($"{DateTime.Now:T}  {e.Folder}: {e.EventTypes}");
                    notified.Set();
                };

                // Applies the filter settings above to the server-side subscription.
                client.UpdateSubscription();
                Console.WriteLine($"{DateTime.Now:T}  watching the Inbox");

                var uri = client.AppendMessage(client.MailboxInfo.InboxUri,
                    new MailMessage("sender@example.com", "receiver@example.com", "Notification test", "Body"));

                try
                {
                    if (!notified.Wait(TimeSpan.FromSeconds(60)))
                        Console.WriteLine($"{DateTime.Now:T}  no notification within 60 seconds");
                }
                finally
                {
                    client.DeleteItem(uri, DeletionOptions.DeletePermanently);
                    client.ResetSubscription();
                }
            }
        }
    }
}
