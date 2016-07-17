
Imports System.Collections.Generic
Imports System.Text
Imports Aspose.Email.Outlook.Pst
Imports Aspose.Email.Outlook

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class UpdateBulkMessagesInPSTFile
        Public Shared Sub Run()
            ' ExStart:UpdateBulkMessagesInPSTFile
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook() + "Sub.pst"

            ' Load the Outlook PST file
            Dim personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir)

            ' Get Requierd Subfolder 
            Dim inbox As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")

            ' find messages having From = "someuser@domain.com"
            Dim queryBuilder As New PersonalStorageQueryBuilder()
            queryBuilder.From.Contains("someuser@domain.com")

            ' Get Contents from Query
            Dim messages As MessageInfoCollection = inbox.GetContents(queryBuilder.GetQuery())

            ' Save (MessageInfo,EntryIdString) in List
            Dim changeList As IList(Of String) = New List(Of String)()
            For Each messageInfo As MessageInfo In messages
                changeList.Add(messageInfo.EntryIdString)
            Next

            ' Compose the new properties
            Dim updatedProperties As New MapiPropertyCollection()
            updatedProperties.Add(MapiPropertyTag.PR_SUBJECT_W, New MapiProperty(MapiPropertyTag.PR_SUBJECT_W, Encoding.Unicode.GetBytes("New Subject")))
            updatedProperties.Add(MapiPropertyTag.PR_IMPORTANCE, New MapiProperty(MapiPropertyTag.PR_IMPORTANCE, BitConverter.GetBytes(CLng(2))))

            ' update messages having From = "someuser@domain.com" with new properties
            inbox.ChangeMessages(changeList, updatedProperties)
        End Sub
        ' ExEnd:UpdateBulkMessagesInPSTFile
    End Class
End Namespace
