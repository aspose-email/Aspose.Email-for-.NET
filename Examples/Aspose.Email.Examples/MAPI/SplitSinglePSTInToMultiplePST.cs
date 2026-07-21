// Demonstrates how to split one large PST into several smaller ones of a given
// size, tracking the progress through the storage events.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SplitSinglePSTInToMultiplePST
    {
        private static int messageCount;
        private static string currentFolder;
        private static int chunks;

        public static void Run()
        {
            messageCount = 0;
            currentFolder = null;
            chunks = 0;

            var outputDir = Data.OutSub("Chunks");

            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                // Subscribing to the events is optional - it only drives the progress
                // output below.
                personalStorage.StorageProcessed += OnStorageProcessed;
                personalStorage.ItemMoved += OnItemMoved;

                // Split into chunks of roughly 5 MB each.
                personalStorage.SplitInto(5000000, outputDir);
            }

            Console.WriteLine($"\nSplit into {chunks} chunk(s) in {outputDir}");
        }

        // Raised once per produced chunk, after its messages have been moved.
        private static void OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            if (currentFolder != null)
                Console.WriteLine($"    Added {messageCount} messages to \"{currentFolder}\"");

            messageCount = 0;
            currentFolder = null;
            chunks++;

            Console.WriteLine($"*** The chunk is processed: {e.FileName}");
        }

        // Raised per moved message; the folder only changes between groups, so print a
        // running total whenever it does.
        private static void OnItemMoved(object sender, ItemMovedEventArgs e)
        {
            var folderPath = e.DestinationFolder.RetrieveFullPath();
            currentFolder = currentFolder ?? folderPath;

            if (currentFolder != folderPath)
            {
                Console.WriteLine($"    Added {messageCount} messages to \"{currentFolder}\"");
                messageCount = 0;
                currentFolder = folderPath;
            }

            messageCount++;
        }
    }
}
