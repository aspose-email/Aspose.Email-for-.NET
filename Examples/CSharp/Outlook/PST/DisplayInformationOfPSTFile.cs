using System;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class DisplayInformationOfPSTFile
    {
        public static void Run()
        {
            // ExStart:DisplayInformationOfPSTFile
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "PersonalStorage.pst";

            // load PST file
            PersonalStorage personalStorage = PersonalStorage.FromFile(dst);

            // Get the folders information
            FolderInfoCollection folderInfoCollection = personalStorage.RootFolder.GetSubFolders();

            // Browse through each folder to display folder name and number of messages
            foreach (FolderInfo folderInfo in folderInfoCollection)
            {
                Console.WriteLine("Folder: " + folderInfo.DisplayName);
                Console.WriteLine("Total items: " + folderInfo.ContentCount);
                Console.WriteLine("Total unread items: " + folderInfo.ContentUnreadCount);
                Console.WriteLine("-----------------------------------");
            }
            // ExEnd:DisplayInformationOfPSTFile
        }
    }
}
