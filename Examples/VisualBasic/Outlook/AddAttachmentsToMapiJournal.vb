Imports System.IO
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddAttachmentsToMapiJournal
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:CreateNewMapiContactAndAddToContactsSubfolder
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim attachFileNames As String() = New String() {dataDir & Convert.ToString("Desert.jpg"), dataDir & Convert.ToString("download.png")}

            Dim journal As New MapiJournal("testJournal", "This is a test journal", "Phone call", "Phone call")
            journal.StartTime = DateTime.Now
            journal.EndTime = journal.StartTime.AddHours(1)
            journal.Companies = New String() {"company 1", "company 2", "company 3"}

            For Each attach As String In attachFileNames
                journal.Attachments.Add(attach, File.ReadAllBytes(attach))
            Next
            journal.Save(dataDir & Convert.ToString("AddAttachmentsToMapiJournal_out.msg"))
            ' ExEnd:AddAttachmentsToMapiJournal
        End Sub
    End Class
End Namespace
