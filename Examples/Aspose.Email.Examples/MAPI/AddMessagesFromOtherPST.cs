// Demonstrates how to copy messages from one PST folder into another PST folder with event notification.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddMessagesFromOtherPst
    {
        private static int added;

        public static void Run()
        {
            added = 0;

            // Work on a copy: messages are added to the destination PST, and examples
            // must never modify the shared input data.
            var destPath = Data.Out/"AddMessagesFromOtherPst_out.pst";
            File.Copy(Data.Mapi/"PersonalStorageFile1.pst", destPath, true);

            // The source is only read, so it can be opened read-only in place.
            using (var srcPst = PersonalStorage.FromFile(Data.Mapi/"SampleContacts.pst", false))
            using (var destPst = PersonalStorage.FromFile(destPath))
            {
                var srcFolder = srcPst.RootFolder.GetSubFolder("Contacts");
                var destFolder = destPst.RootFolder.GetSubFolder("myInbox");

                destFolder.MessageAdded += OnMessageAdded;
                destFolder.AddMessages(srcFolder.EnumerateMapiMessages());
            }

            Console.WriteLine($"\nAdded {added} message(s) to {destPath}");
        }

        private static void OnMessageAdded(object sender, MessageAddedEventArgs e)
        {
            added++;
            Console.WriteLine($"Added: {e.EntryId}");
        }
    }
}
