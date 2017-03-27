using System;
using System.IO;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SplitSinglePSTInToMultiplePST
    {
        public int totalAdded;
        public static int messageCount;
        public static string currentFolder;
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:SplitSinglePSTInToMultiplePST
            string dataDir = RunExamples.GetDataDir_Outlook();
            try
            {
                String dstSplit = dataDir + Convert.ToString("Chunks\\");

                // Delete the files if already present
                foreach (string file__1 in Directory.GetFiles(dstSplit))
                {
                    File.Delete(file__1);
                }

                using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "Sub.pst"))
                {
                    // The events subscription is an optional step for the tracking process only.
                    personalStorage.StorageProcessed += PstSplit_OnStorageProcessed;
                    personalStorage.ItemMoved += PstSplit_OnItemMoved;

                    // Splits into pst chunks with the size of 5mb
                    personalStorage.SplitInto(5000000, dataDir + @"\Chunks\");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\nThis example will only work if you apply a valid Aspose Email License. You can purchase full license or get 30 day temporary license from http:// Www.aspose.com/purchase/default.aspx.");
            }
            // ExEnd:SplitSinglePSTInToMultiplePST
        }

        void destinationFolder_ItemMoved(object sender, ItemMovedEventArgs e)
        {
            totalAdded++;
        }

        void PstMerge_OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            Console.WriteLine("*** The storage is merging: {0}", e.FileName);
        }

        void PstMerge_OnItemMoved(object sender, ItemMovedEventArgs e)
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
    }
}