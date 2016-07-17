using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Outlook.Pst;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SplitPST
    {
        private static string currentFolder;
        private static int messageCount;

        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "Outlook.pst";
            String dstSplit = dataDir + @"Chunks-Split\";

            // Delete the files if already present
            foreach (string file in Directory.GetFiles(dstSplit))
                File.Delete(file);

            using (PersonalStorage pst = PersonalStorage.FromFile(dst))
            {
                // The events subscription is an optional step for the tracking process only.
                pst.StorageProcessed += PstSplit_OnStorageProcessed;
                pst.ItemMoved += PstSplit_OnItemMoved;

                // Splits into pst chunks with the size of 300 KB
                pst.SplitInto(300 * 1024, dstSplit);
            }

            Console.WriteLine(Environment.NewLine + "PST split successfully at " + dst);
        }

        static void PstSplit_OnItemMoved(object sender, ItemMovedEventArgs e)
        {
            if (currentFolder == null)
            {
                currentFolder = e.DestinationFolder.RetrieveFullPath();
            }

            string folderPath = e.DestinationFolder.RetrieveFullPath();

            if (currentFolder != folderPath)
            {
                Console.WriteLine("    Added {0} messages to \"{1}\"", messageCount, currentFolder);
                messageCount = 0;
                currentFolder = folderPath;
            }

            messageCount++;
        }

        static void PstSplit_OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            if (currentFolder != null)
            {
                Console.WriteLine("    Added {0} messages to \"{1}\"", messageCount, currentFolder);
            }

            messageCount = 0;
            currentFolder = null;
            Console.WriteLine("*** The chunk is processed: {0}", e.FileName);
        }
    }
}
