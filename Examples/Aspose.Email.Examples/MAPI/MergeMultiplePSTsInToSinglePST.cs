// Demonstrates how to merge multiple PST files into one PST with progress tracking.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class MergeMultiplePstsInToSinglePst
    {
        private static int totalAdded;
        private static string currentFolder;
        private static int messageCount;

        public static void Run()
        {
            totalAdded = 0;
            currentFolder = null;
            messageCount = 0;

            // Work on a copy: the merge writes into the destination PST, and examples
            // must never modify the shared input data.
            var pstPath = Data.Out/"MergeMultiplePstsInToSinglePst_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var personalStorage = PersonalStorage.FromFile(pstPath))
            {
                personalStorage.StorageProcessed += OnStorageProcessed;
                personalStorage.ItemMoved += OnItemMoved;
                personalStorage.MergeWith(Directory.GetFiles(Data.Mapi/"MergePST"));

                Console.WriteLine($"    Added {messageCount} messages to \"{currentFolder}\"");
                Console.WriteLine($"\nTotal messages added: {totalAdded}");
                Console.WriteLine($"Merged into {pstPath}");
            }
        }

        private static void OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            Console.WriteLine("*** The storage is merging: {0}", e.FileName);
        }

        private static void OnItemMoved(object sender, ItemMovedEventArgs e)
        {
            if (currentFolder == null)
                currentFolder = e.DestinationFolder.RetrieveFullPath();

            string folderPath = e.DestinationFolder.RetrieveFullPath();
            if (currentFolder != folderPath)
            {
                Console.WriteLine("    Added {0} messages to \"{1}\"", messageCount, currentFolder);
                messageCount = 0;
                currentFolder = folderPath;
            }
            messageCount++;
            totalAdded++;
        }
    }
}
