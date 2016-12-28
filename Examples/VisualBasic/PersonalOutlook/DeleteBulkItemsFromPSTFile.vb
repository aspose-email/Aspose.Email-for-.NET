Imports System.Collections.Generic
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class DeleteBulkItemsFromPSTFile
        Public Shared Sub Run()
            ' ExStart:DeleteBulkItemsFromPSTFile
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook() + "Sub.pst"

            Using personalStorage1 As PersonalStorage = PersonalStorage.FromFile(dataDir)
                ' Get Inbox SubFolder from Outlook file
                Dim inbox As FolderInfo = personalStorage1.RootFolder.GetSubFolder("Inbox")

                ' Create instance of PersonalStorageQueryBuilder
                Dim queryBuilder As New PersonalStorageQueryBuilder()

                queryBuilder.From.Contains("someuser@domain.com")
                Dim messages As MessageInfoCollection = inbox.GetContents(queryBuilder.GetQuery())
                Dim deleteList As IList(Of String) = New List(Of String)()
                For Each messageInfo As MessageInfo In messages
                    deleteList.Add(messageInfo.EntryIdString)
                Next

                ' delete messages having From = "someuser@domain.com"
                inbox.DeleteChildItems(deleteList)
            End Using
            ' ExEnd:DeleteBulkItemsFromPSTFile
        End Sub
    End Class
End Namespace