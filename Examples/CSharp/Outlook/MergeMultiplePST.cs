using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Outlook.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class MergeMultiplePST
    {
        public static int totalAdded=0;
        public static int messageCount;
        public static string currentFolder;
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:MergeMultiplePST
            string dataDir = RunExamples.GetDataDir_Outlook();
            using (PersonalStorage pst = PersonalStorage.FromFile(dataDir + "Test.pst"))
            {
                // The events subscription is an optional step for the tracking process only.
                pst.StorageProcessed += PstMerge_OnStorageProcessed;
                pst.ItemMoved += PstMerge_OnItemMoved;

                // Merges with the pst files that are located in separate folder.
                pst.MergeWith(Directory.GetFiles(dataDir + @"\Sources\"));
                Console.WriteLine("Total messages added: {0}", totalAdded);
            }
            // ExEnd:MergeMultiplePST
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