Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ChangeFolderContainerClass
        Public Shared Sub Run()
            ' ExStart:ChangeFolderContainerClass
            ' The path to the file directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook file
            Dim path As String = dataDir & Convert.ToString("PersonalStorage1.pst")

            ' Create new PST
            Dim personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(path)

            Dim folder As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")
            folder.ChangeContainerClass("IPF.Note")
            personalStorage__1.Dispose()
            ' ExEnd:ChangeFolderContainerClass
        End Sub
    End Class
End Namespace
