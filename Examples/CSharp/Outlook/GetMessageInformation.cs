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
    class GetMessageInformation
    {
        // ExStart:GetMessageInformation
        public static void Run()
        {            
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "PersonalStorage.pst";

            try
            {
               
                // Load the Outlook PST file
                PersonalStorage personalStorage = PersonalStorage.FromFile(path);

                // Get the Display Format of the PST file
                Console.WriteLine("Display Format: " + personalStorage.Format);

                // Get the folders and messages information
                FolderInfo folderInfo = personalStorage.RootFolder;

                // Call the recursive method to display the folder contents
                DisplayFolderContents(folderInfo, personalStorage);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);               
            }            
        }

        /// <summary>
        /// This is a recursive method to display contents of a folder
        /// </summary>
        /// <param name="folderInfo"></param>
        /// <param name="pst"></param>
        private static void DisplayFolderContents(FolderInfo folderInfo, PersonalStorage pst)
        {
            // Display the folder name
            Console.WriteLine("Folder: " + folderInfo.DisplayName);
            Console.WriteLine("==================================");
            // Display information about messages inside this folder
            MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
            foreach (MessageInfo messageInfo in messageInfoCollection)
            {
                Console.WriteLine("Subject: " + messageInfo.Subject);
                Console.WriteLine("Sender: " + messageInfo.SenderRepresentativeName);
                Console.WriteLine("Recipients: " + messageInfo.DisplayTo);
                Console.WriteLine("------------------------------");
            }

            // Call this method recursively for each subfolder
            if (folderInfo.HasSubFolders == true)
            {
                foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
                {
                    DisplayFolderContents(subfolderInfo, pst);
                }
            }
        }
        // ExEnd:GetMessageInformation
    }
}
