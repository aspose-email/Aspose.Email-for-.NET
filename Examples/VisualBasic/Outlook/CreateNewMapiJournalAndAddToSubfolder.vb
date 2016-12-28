
Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateNewMapiJournalAndAddToSubfolder
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:CreateNewMapiJournalAndAddToSubfolder
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim journal As New MapiJournal("daily record", "called out in the dark", "Phone call", "Phone call")
            journal.StartTime = DateTime.Now
            journal.EndTime = journal.StartTime.AddHours(1)

            Dim checkfile As String = dataDir + "CreateNewMapiJournalAndAddToSubfolder_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If

            Using personalStorage__1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("CreateNewMapiJournalAndAddToSubfolder_out.pst"), FileFormatVersion.Unicode)
                Dim journalFolder As FolderInfo = personalStorage__1.CreatePredefinedFolder("Journal", StandardIpmFolder.Journal)
                journalFolder.AddMapiMessageItem(journal)
            End Using
            ' ExEnd:CreateNewMapiJournalAndAddToSubfolder
        End Sub
    End Class
End Namespace
