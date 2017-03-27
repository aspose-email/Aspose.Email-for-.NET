using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email.Storage.Pst;
using Aspose.Email;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class DeleteMessagesFromPSTFiles
    {
        public static void Run()
        {
            // ExStart:DeleteMessagesFromPSTFiles
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook() + "Sub.pst";

            // Load the Outlook PST file
            PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir);

            // Get the Sent items folder
            FolderInfo folderInfo = personalStorage.GetPredefinedFolder(StandardIpmFolder.SentItems);

            MessageInfoCollection msgInfoColl = folderInfo.GetContents();
            foreach (MessageInfo msgInfo in msgInfoColl)
            {
                Console.WriteLine(msgInfo.Subject + ": " + msgInfo.EntryIdString);
                if (msgInfo.Subject.Equals("some delete condition") == true)
                {
                    // Delete this item
                    folderInfo.DeleteChildItem(msgInfo.EntryId);
                    Console.WriteLine("Deleted this message");
                }
            }
            // ExEnd:DeleteMessagesFromPSTFiles
        }
    }
}
