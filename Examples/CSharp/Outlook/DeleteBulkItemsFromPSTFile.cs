using System.Collections.Generic;
using System.IO;
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
    class DeleteBulkItemsFromPSTFile
    {
        public static void Run()
        {
            // ExStart:DeleteBulkItemsFromPSTFile
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook() + @"Sub.pst";
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir))
            {
                // Get Inbox SubFolder from Outlook file
                FolderInfo inbox = personalStorage.RootFolder.GetSubFolder("Inbox");

                // Create instance of PersonalStorageQueryBuilder
                PersonalStorageQueryBuilder queryBuilder = new PersonalStorageQueryBuilder();

                queryBuilder.From.Contains("someuser@domain.com");
                MessageInfoCollection messages = inbox.GetContents(queryBuilder.GetQuery());
                IList<string> deleteList = new List<string>();
                foreach (MessageInfo messageInfo in messages)
                {
                    deleteList.Add(messageInfo.EntryIdString);
                }

                // delete messages having From = "someuser@domain.com"
                inbox.DeleteChildItems(deleteList);
            }
            // ExEnd:DeleteBulkItemsFromPSTFile
        }
    }
}