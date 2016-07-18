Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class RetreiveParentFolderInformationFromMessageInfo
        Public Shared Sub Run()
            ' ExStart:RetreiveParentFolderInformationFromMessageInfo
            ' Load Pst File
            Dim dataDir As String = RunExamples.GetDataDir_Outlook() + "PersonalStorage.pst"
            Using personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir)
                For Each folder As FolderInfo In personalStorage__1.RootFolder.GetSubFolders()
                    For Each msg As MessageInfo In folder.EnumerateMessages()
                        Dim fi As FolderInfo = personalStorage__1.GetParentFolder(msg.EntryId)
                    Next
                Next
            End Using
            ' ExEnd:RetreiveParentFolderInformationFromMessageInfo
        End Sub
    End Class
End Namespace
