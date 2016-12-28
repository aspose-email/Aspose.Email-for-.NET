Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AccessContactInformation
        Public Shared Sub Run()
            ' ExStart:AccessContactInformation
            ' Load the Outlook file
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook PST file
            Dim personalStorage1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("SampleContacts.pst"))
            ' Get the Contacts folder
            Dim folderInfo As FolderInfo = personalStorage1.RootFolder.GetSubFolder("Contacts")
            ' Loop through all the contacts in this folder
            Dim messageInfoCollection As MessageInfoCollection = folderInfo.GetContents()
            For Each messageInfo As MessageInfo In messageInfoCollection
                ' Get the contact information
                Dim mapi As MapiMessage = personalStorage1.ExtractMessage(messageInfo)

                Dim contact As MapiContact = DirectCast(mapi.ToMapiMessageItem(), MapiContact)

                ' Display some contents on screen
                Console.WriteLine("Name: " + contact.NameInfo.DisplayName)
                ' Save to disk in MSG format
                If contact.NameInfo.DisplayName IsNot Nothing Then
                    Dim message As MapiMessage = personalStorage1.ExtractMessage(messageInfo)
                    ' Get rid of illegal characters that cannot be used as a file name
                    Dim messageName As String = message.Subject.Replace(":", " ").Replace("\", " ").Replace("?", " ").Replace("/", " ")
                    message.Save((Convert.ToString(dataDir & Convert.ToString("Contacts\")) & messageName) + "_out.msg")
                End If
            Next
            ' ExEnd:AccessContactInformation
        End Sub
    End Class
End Namespace
