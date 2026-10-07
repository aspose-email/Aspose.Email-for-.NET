// Demonstrates being notified about new messages as they arrive (IMAP IDLE).
//
// StartMonitoring keeps watching a folder and calls the callback whenever messages are
// added or removed, so there is no need to poll. Here the change is made by the example
// itself: it appends a message over a second connection while the first one waits.
// ResumeMonitoringAfterError shows how to recover when monitoring stops on an error.
// The example watches a uniquely named folder and deletes it at the end.

using System;
using System.Threading;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SupportIMAPIdleCommand
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            using (var changed = new ManualResetEventSlim(false))
            {
                client.CreateFolder(folderName);

                try
                {
                    ImapMonitoringEventHandler onChange = (sender, e) =>
                    {
                        foreach (var info in e.NewMessages)
                            Console.WriteLine($"{DateTime.Now:T}  new in {e.FolderName}: {info.Subject}");
                        if (e.DeletedMessages.Length > 0)
                            Console.WriteLine($"{DateTime.Now:T}  {e.DeletedMessages.Length} message(s) removed");
                        changed.Set();
                    };

                    ImapMonitoringErrorEventHandler onError = (sender, e) =>
                    {
                        Console.WriteLine($"{DateTime.Now:T}  monitoring stopped: {e.Error.Message}");
                        changed.Set();
                    };

                    client.StartMonitoring(onChange, onError, folderName);
                    Console.WriteLine($"{DateTime.Now:T}  watching '{folderName}'");

                    using (var connection = client.CreateConnection())
                    {
                        client.AppendMessage(connection, folderName,
                            new MailMessage("sender@example.com", "receiver@example.com", "Arrived while idle", "Body"));
                        Console.WriteLine($"{DateTime.Now:T}  appended a message over a second connection");
                    }

                    if (!changed.Wait(TimeSpan.FromSeconds(30)))
                        Console.WriteLine($"{DateTime.Now:T}  no notification within 30 seconds");

                    client.StopMonitoring(folderName);
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
