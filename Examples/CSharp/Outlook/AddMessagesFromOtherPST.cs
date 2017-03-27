using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Email;
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
    class AddMessagesFromOtherPST
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "SampleContacts.pst";
            BulkAddFromAnotherPst(path);
        }

        // ExStart:AddMessagesFromOtherPST
        private static void BulkAddFromAnotherPst(string source)
        {
            using (PersonalStorage pst = PersonalStorage.FromFile(source, false))
            using (PersonalStorage pstDest = PersonalStorage.FromFile(RunExamples.GetDataDir_Outlook() + "PersonalStorageFile1.pst"))
            {
                // Get the folder by name
                FolderInfo folderInfo = pst.RootFolder.GetSubFolder("Contacts");
                MessageInfoCollection ms = folderInfo.GetContents();

                // Get the folder by name
                FolderInfo f = pstDest.RootFolder.GetSubFolder("myInbox");
                f.MessageAdded += new MessageAddedEventHandler(OnMessageAdded);
                f.AddMessages(folderInfo.EnumerateMapiMessages());
                FolderInfo fi = pstDest.RootFolder.GetSubFolder("myInbox");
                MessageInfoCollection msgs = fi.GetContents();

            }

        }

        /// <summary>
        /// Handles the MessageAdded event.
        /// </summary>
        static void OnMessageAdded(object sender, MessageAddedEventArgs e)
        {
            Console.WriteLine(e.EntryId);
            Console.WriteLine(e.Message.Subject);
        }
        // ExEnd:AddMessagesFromOtherPST
    }
}