// Demonstrates how to bulk-add MSG files to a PST folder using AddMessages with event notification.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddingBulkMessagesWithImprovedPerformance
    {
        public static void Run()
        {
            File.Copy(Data.Mapi/"PersonalStorageFile2.pst", Data.Out/"test.pst", overwrite: true);

            using (var pst = PersonalStorage.FromFile(Data.Out/"test.pst"))
            {
                var folder = pst.RootFolder.GetSubFolder("myInbox");
                folder.MessageAdded += OnMessageAdded;
                folder.AddMessages(GetMessages(Data.Mapi/"Msg"));
            }
        }

        private static IEnumerable<MapiMessage> GetMessages(string dir)
        {
            foreach (var file in Directory.GetFiles(dir, "*.msg"))
                yield return MapiMessage.Load(file);
        }

        private static void OnMessageAdded(object sender, MessageAddedEventArgs e)
        {
            Console.WriteLine($"Added: {e.EntryId}");
        }
    }
}
