// Demonstrates stopping a PST restore part-way.
//
// RestoreSettings.BeforeItemCallback is called before every folder and item is
// uploaded. Throwing an exception from it aborts the restore - useful to enforce a time
// limit, react to a cancel button, or upload only part of a large PST. The example
// builds a PST with 20 messages in a uniquely named folder, stops after five items,
// and removes the restored folder at the end.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.EWS
{
    internal static class AbortPSTToExchangeServerOperation
    {
        private sealed class RestoreAbortedException : Exception
        {
        }

        public static void Run()
        {
            const int maxItems = 5;
            var folderName = "Aspose-Restore-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var stream = new MemoryStream())
            using (var pst = PersonalStorage.Create(stream, FileFormatVersion.Unicode))
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = pst.RootFolder.AddSubFolder(folderName);
                for (var i = 1; i <= 20; i++)
                    folder.AddMessage(new MapiMessage("from@example.com", "to@example.com", $"Message {i}", "Body"));

                var processed = 0;
                var settings = new RestoreSettings
                {
                    BeforeItemCallback = item =>
                    {
                        if (processed == maxItems)
                            throw new RestoreAbortedException();
                        processed++;
                    }
                };

                try
                {
                    client.Restore(pst, settings);
                    Console.WriteLine("The restore finished before the limit was reached.");
                }
                catch (RestoreAbortedException)
                {
                    Console.WriteLine($"Restore stopped after {processed} item(s).");
                }
                finally
                {
                    var restored = client.ListSubFolders(client.MailboxInfo.RootUri)
                        .FirstOrDefault(f => f.DisplayName == folderName);
                    if (restored != null)
                    {
                        Console.WriteLine($"'{folderName}' received {restored.TotalCount} message(s); deleting it.");
                        client.DeleteFolder(restored.Uri, true);
                    }
                }
            }
        }
    }
}
