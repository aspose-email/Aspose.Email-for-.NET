using System;
using System.IO;
using Aspose.Email.Outlook.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class MergeFolderFromAnotherPSTFile
    {
        public static int totalAdded = 0;
        public static int messageCount;
        public static string currentFolder;
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:MergeMultiplePST
            string dataDir = RunExamples.GetDataDir_Outlook();
            try
            {
                using (PersonalStorage pst = PersonalStorage.FromFile(dataDir + "PersonalStorage.pst"))
                {
                    // The events subscription is an optional step for the tracking process only.
                    pst.StorageProcessed += PstMerge_OnStorageProcessed;
                    pst.ItemMoved += PstMerge_OnItemMoved;

                    // Merges with the pst files that are located in separate folder.
                    pst.MergeWith(Directory.GetFiles(dataDir + @"\Sources\"));
                    Console.WriteLine("Total messages added: {0}", totalAdded);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\nThis example will only work if you apply a valid Aspose Email License. You can purchase full license or get 30 day temporary license from http:// Www.aspose.com/purchase/default.aspx.");
            }
            // ExEnd:MergeMultiplePST
        }

        static void PstMerge_OnStorageProcessed(object sender, StorageProcessedEventArgs e)
        {
            Console.WriteLine("*** The storage is merging: {0}", e.FileName);
        }

        static void PstMerge_OnItemMoved(object sender, ItemMovedEventArgs e)
        {
            string folderPath = e.DestinationFolder.RetrieveFullPath();

            if (currentFolder == null)
            {
                currentFolder = e.DestinationFolder.RetrieveFullPath();
            }

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