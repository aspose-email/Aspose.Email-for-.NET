Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
' * 
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SaveContactInformation
        Public Shared Sub Run()
            ' ExStart:SaveContactInformation
            ' Load the Outlook file
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook PST file
            Dim personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir + "Outlook.pst")
            ' Get the Contacts folder
            Dim folderInfo As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Contacts")
            ' Loop through all the contacts in this folder
            Dim messageInfoCollection As MessageInfoCollection = folderInfo.GetContents()
            For Each messageInfo As MessageInfo In messageInfoCollection
                ' Get the contact information
                Dim contact As MapiContact = DirectCast(personalStorage__1.ExtractMessage(messageInfo).ToMapiMessageItem(), MapiContact)
                ' Display some contents on screen
                Console.WriteLine("Name: " + contact.NameInfo.DisplayName + " - " + messageInfo.EntryIdString)
                ' Save to disk in vCard VCF format
                contact.Save(dataDir + "Contacts\\" + contact.NameInfo.DisplayName + ".vcf", ContactSaveFormat.VCard)
            Next
            ' ExEnd:SaveContactInformation
        End Sub
    End Class
End Namespace