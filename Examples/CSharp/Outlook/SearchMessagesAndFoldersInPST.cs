using System;
using Aspose.Email.Mapi;
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
    class SearchMessagesAndFoldersInPST
    {
        public static void Run()
        {
            
            // ExStart:SearchMessagesAndFoldersInPST
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "Outlook.pst"))
            {
                FolderInfo folder = personalStorage.RootFolder.GetSubFolder("Inbox");
                PersonalStorageQueryBuilder builder = new PersonalStorageQueryBuilder();

                // High importance messages
                builder.Importance.Equals((int)MapiImportance.High);
                MessageInfoCollection messages = folder.GetContents(builder.GetQuery());
                Console.WriteLine("Messages with High Imp:" + messages.Count);

                builder = new PersonalStorageQueryBuilder();
                builder.MessageClass.Equals("IPM.Note");
                messages = folder.GetContents(builder.GetQuery());
                Console.WriteLine("Messages with IPM.Note:" + messages.Count);

                builder = new PersonalStorageQueryBuilder();
                // Messages with attachments AND high importance
                builder.Importance.Equals((int)MapiImportance.High);
                builder.HasFlags(MapiMessageFlags.MSGFLAG_HASATTACH);
                messages = folder.GetContents(builder.GetQuery());
                Console.WriteLine("Messages with atts: " + messages.Count);

                builder = new PersonalStorageQueryBuilder();
                // Messages with size > 15 KB
                builder.MessageSize.Greater(15000);
                messages = folder.GetContents(builder.GetQuery());
                Console.WriteLine("messags size > 15Kb:" + messages.Count);

                builder = new PersonalStorageQueryBuilder();
                // Unread messages
                builder.HasNoFlags(MapiMessageFlags.MSGFLAG_READ);
                messages = folder.GetContents(builder.GetQuery());
                Console.WriteLine("Unread:" + messages.Count);

                builder = new PersonalStorageQueryBuilder();
                // Unread messages with attachments
                builder.HasNoFlags(MapiMessageFlags.MSGFLAG_READ);
                builder.HasFlags(MapiMessageFlags.MSGFLAG_HASATTACH);
                messages = folder.GetContents(builder.GetQuery());
                Console.WriteLine("Unread msgs with atts: " + messages.Count);

                // Folder with name of 'SubInbox'
                builder = new PersonalStorageQueryBuilder();
                builder.FolderName.Equals("SubInbox");
                FolderInfoCollection folders = folder.GetSubFolders(builder.GetQuery());
                Console.WriteLine("Folder having subfolder: " + folders.Count);

                builder = new PersonalStorageQueryBuilder();
                // Folders with subfolders
                builder.HasSubfolders();
                folders = folder.GetSubFolders(builder.GetQuery());
                Console.WriteLine(folders.Count);
            }
            // ExEnd:SearchMessagesAndFoldersInPST
        }
    }
}