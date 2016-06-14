using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Outlook.Pst;

namespace Aspose.Email.Examples.CSharp.Outlook
{
    class MergePSTFiles
    {
        static int totalAdded;
        static string currentFolder;
        static int messageCount;

        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "Test.pst";

            totalAdded = 0;

            using (PersonalStorage pst = PersonalStorage.FromFile(dst))
            {
                // The events subscription is an optional step for the tracking process only.
                pst.StorageProcessed += PstMerge_OnStorageProcessed;
                pst.ItemMoved += PstMerge_OnItemMoved;

                // Merges with the pst files that are located in separate folder. 
                pst.MergeWith(Directory.GetFiles(dataDir + @"chunks\"));
                Console.WriteLine("Total messages added: {0}", totalAdded);
            }

            Console.WriteLine(Environment.NewLine + "PST merged successfully at " + dst);
        }

        static void PstMerge_OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            Console.WriteLine("*** The storage is merging: {0}", e.FileName);
        }

        static void PstMerge_OnItemMoved(object sender, ItemMovedEventArgs e)
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
            totalAdded++;
        }
    }
}
