Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ExtractNumberOfMessages
        Public Shared Sub Run()
            ' ExStart:ExtractNumberOfMessages
            ' The path to the file directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook file
            Dim path As String = dataDir & Convert.ToString("Sub.pst")

            ' Save message to file
            Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(path)
                Dim inbox As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")
                ' Extracts messages starting from 10th index top and extract total 100 messages
                Dim messages As MessageInfoCollection = inbox.GetContents(10, 100)
            End Using
            ' ExEnd:ExtractNumberOfMessages
        End Sub
    End Class
End Namespace
