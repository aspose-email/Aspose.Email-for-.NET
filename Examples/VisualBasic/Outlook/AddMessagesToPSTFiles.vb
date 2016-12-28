Imports System.IO
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
    Class AddMessagesToPSTFiles
        Public Shared Sub Run()
            ' ExStart:AddMessagesToPSTFiles
            ' The path to the file directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim checkfile As String = dataDir + "AddMessagesToPSTFiles_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)

            Else
            End If

            ' Create new PST            
            Dim personalStorage__1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("AddMessagesToPSTFiles_out.pst"), FileFormatVersion.Unicode)

            ' Add new folder "Inbox"
            personalStorage__1.RootFolder.AddSubFolder("Inbox")

            ' Select the "Inbox" folder
            Dim inboxFolder As FolderInfo = personalStorage__1.RootFolder.GetSubFolder("Inbox")

            ' Add some messages to "Inbox" folder
            inboxFolder.AddMessage(MapiMessage.FromFile(dataDir & Convert.ToString("MapiMsgWithPoll.msg")))
            ' ExEnd:AddMessagesToPSTFiles            
        End Sub
    End Class
End Namespace