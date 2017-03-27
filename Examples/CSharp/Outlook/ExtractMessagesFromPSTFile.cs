using System.IO;
using System;
using Aspose.Email.Mapi;
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
    class ExtractMessagesFromPSTFile
    {
        // ExStart:ExtractMessagesFromPSTFile
        public static void Run()
        {            
            // The path to the file directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "PersonalStorage.pst";

            try
            {
                // load the Outlook PST file
                PersonalStorage pst = PersonalStorage.FromFile(path);

                // get the Display Format of the PST file
                Console.WriteLine("Display Format: " + pst.Format);

                // get the folders and messages information
                FolderInfo folderInfo = pst.RootFolder;

                // Call the recursive method to extract msg files from each folder
                ExtractMsgFiles(folderInfo, pst);
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
        private static void ExtractMsgFiles(FolderInfo folderInfo, PersonalStorage pst)
        {
            // display the folder name
            Console.WriteLine("Folder: " + folderInfo.DisplayName);
            Console.WriteLine("==================================");
            // loop through all the messages in this folder
            MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
            foreach (MessageInfo messageInfo in messageInfoCollection)
            {
                Console.WriteLine("Saving message {0} ....", messageInfo.Subject);
                // get the message in MapiMessage instance
                MapiMessage message = pst.ExtractMessage(messageInfo);
                // save this message to disk in msg format
                message.Save(message.Subject.Replace(":", " ") + ".msg");
                // save this message to stream in msg format
                MemoryStream messageStream = new MemoryStream();
                message.Save(messageStream);
            }

            // Call this method recursively for each subfolder
            if (folderInfo.HasSubFolders == true)
            {
                foreach (FolderInfo subfolderInfo in folderInfo.GetSubFolders())
                {
                    ExtractMsgFiles(subfolderInfo, pst);
                }
            }
        }
        // ExEnd:ExtractMessagesFromPSTFile
    }
}
