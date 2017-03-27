using System.IO;
using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class MergeMultiplePSTsInToSinglePST
    {
        static int totalAdded;
        static string currentFolder;
        static int messageCount;

        public static void Run()
        {
            // ExStart:MergePSTFiles
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "Sub.pst";
            totalAdded = 0;
            try
            {
                using (PersonalStorage personalStorage = PersonalStorage.FromFile(dst))
                {
                    // The events subscription is an optional step for the tracking process only.
                    personalStorage.StorageProcessed += PstMerge_OnStorageProcessed;
                    personalStorage.ItemMoved += PstMerge_OnItemMoved;

                    // Merges with the pst files that are located in separate folder. 
                    personalStorage.MergeWith(Directory.GetFiles(dataDir + @"MergePST\"));
                    Console.WriteLine("Total messages added: {0}", totalAdded);
                }                
                Console.WriteLine(Environment.NewLine + "PST merged successfully at " + dst);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\nThis example will only work if you apply a valid Aspose Email License. You can purchase full license or get 30 day temporary license from http:// Www.aspose.com/purchase/default.aspx.");
            }
            // ExEnd:MergePSTFiles
        }

        // ExStart:HelperMethods
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
        // ExEnd:HelperMethods
    }
}