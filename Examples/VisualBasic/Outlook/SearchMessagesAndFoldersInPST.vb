Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SearchMessagesAndFoldersInPST
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:SearchMessagesAndFoldersInPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Outlook.pst"))
                Dim folder As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")
                Dim builder As New PersonalStorageQueryBuilder()

                ' High importance messages
                builder.Importance.Equals(CInt(MapiImportance.High))
                Dim messages As MessageInfoCollection = folder.GetContents(builder.GetQuery())
                Console.WriteLine("Messages with High Imp:" + messages.Count.ToString())

                builder = New PersonalStorageQueryBuilder()
                builder.MessageClass.Equals("IPM.Note")
                messages = folder.GetContents(builder.GetQuery())
                Console.WriteLine("Messages with IPM.Note:" + messages.Count.ToString())

                builder = New PersonalStorageQueryBuilder()
                ' Messages with attachments AND high importance
                builder.Importance.Equals(CInt(MapiImportance.High))
                builder.HasFlags(MapiMessageFlags.MSGFLAG_HASATTACH)
                messages = folder.GetContents(builder.GetQuery())
                Console.WriteLine("Messages with atts: " + messages.Count.ToString())

                builder = New PersonalStorageQueryBuilder()
                ' Messages with size > 15 KB
                builder.MessageSize.Greater(15000)
                messages = folder.GetContents(builder.GetQuery())
                Console.WriteLine("messags size > 15Kb:" + messages.Count.ToString())

                builder = New PersonalStorageQueryBuilder()
                ' Unread messages
                builder.HasNoFlags(MapiMessageFlags.MSGFLAG_READ)
                messages = folder.GetContents(builder.GetQuery())
                Console.WriteLine("Unread:" + messages.Count.ToString())

                builder = New PersonalStorageQueryBuilder()
                ' Unread messages with attachments
                builder.HasNoFlags(MapiMessageFlags.MSGFLAG_READ)
                builder.HasFlags(MapiMessageFlags.MSGFLAG_HASATTACH)
                messages = folder.GetContents(builder.GetQuery())
                Console.WriteLine("Unread msgs with atts: " + messages.Count.ToString())

                ' Folder with name of 'SubInbox'
                builder = New PersonalStorageQueryBuilder()
                builder.FolderName.Equals("SubInbox")
                Dim folders As FolderInfoCollection = folder.GetSubFolders(builder.GetQuery())
                Console.WriteLine("Folder having subfolder: " + folders.Count.ToString())

                builder = New PersonalStorageQueryBuilder()
                ' Folders with subfolders
                builder.HasSubfolders()
                folders = folder.GetSubFolders(builder.GetQuery())
                Console.WriteLine(folders.Count.ToString())
            End Using
            ' ExEnd:SearchMessagesAndFoldersInPST
        End Sub
    End Class
End Namespace
