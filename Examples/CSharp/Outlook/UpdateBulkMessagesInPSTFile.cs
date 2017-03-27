using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Mapi;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class UpdateBulkMessagesInPSTFile
    {
        public static void Run()
        {
            // ExStart:UpdateBulkMessagesInPSTFile
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook() + "Sub.pst";

            // Load the Outlook PST file
            PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir);
            
            // Get Requierd Subfolder 
            FolderInfo inbox = personalStorage.RootFolder.GetSubFolder("Inbox");

            // find messages having From = "someuser@domain.com"
            PersonalStorageQueryBuilder queryBuilder = new PersonalStorageQueryBuilder();
            queryBuilder.From.Contains("someuser@domain.com");

            // Get Contents from Query
            MessageInfoCollection messages = inbox.GetContents(queryBuilder.GetQuery());

            // Save (MessageInfo,EntryIdString) in List
            IList<string> changeList = new List<string>();
            foreach (MessageInfo messageInfo in messages)
            {
                changeList.Add(messageInfo.EntryIdString);
            }

            // Compose the new properties
            MapiPropertyCollection updatedProperties = new MapiPropertyCollection();
            updatedProperties.Add(MapiPropertyTag.PR_SUBJECT_W, new MapiProperty(MapiPropertyTag.PR_SUBJECT_W, Encoding.Unicode.GetBytes("New Subject")));
            updatedProperties.Add(MapiPropertyTag.PR_IMPORTANCE, new MapiProperty(MapiPropertyTag.PR_IMPORTANCE, BitConverter.GetBytes((long)2)));

            // update messages having From = "someuser@domain.com" with new properties
            inbox.ChangeMessages(changeList, updatedProperties);
            // ExEnd:UpdateBulkMessagesInPSTFile
        }
    }
}
