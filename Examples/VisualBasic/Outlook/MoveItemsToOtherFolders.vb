Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class MoveItemsToOtherFolders
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:MoveItemsToOtherFolders
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Using personalStorage1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Outlook_1.pst"))
                Dim inbox As FolderInfo = personalStorage1.GetPredefinedFolder(StandardIpmFolder.Inbox)
                Dim deleted As FolderInfo = personalStorage1.GetPredefinedFolder(StandardIpmFolder.DeletedItems)
                Dim subfolder As FolderInfo = inbox.GetSubFolder("Inbox")

                If subfolder IsNot Nothing Then
                    ' Move folder to the Deleted Items
                    personalStorage1.MoveItem(subfolder, deleted)

                    ' Move message to the Deleted Items
                    Dim contents As MessageInfoCollection = subfolder.GetContents()
                    personalStorage1.MoveItem(contents(0), deleted)

                    ' Move all inbox subfolders to the Deleted Items
                    inbox.MoveSubfolders(deleted)

                    ' Move all subfolder contents to the Deleted Items
                    subfolder.MoveContents(deleted)
                End If
            End Using
            ' ExEnd:MoveItemsToOtherFolders
        End Sub
    End Class
End Namespace
