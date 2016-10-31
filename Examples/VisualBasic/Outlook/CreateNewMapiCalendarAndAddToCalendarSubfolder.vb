Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateNewMapiCalendarAndAddToCalendarSubfolder
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:CreateNewMapiCalendarAndAddToCalendarSubfolder
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim message As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("Note.msg"))

            ' Note #1
            Dim note1 As MapiNote = DirectCast(message.ToMapiMessageItem(), MapiNote)
            note1.Subject = "Yellow color note"
            note1.Body = "This is a yellow color note"

            ' Note #2
            Dim note2 As MapiNote = DirectCast(message.ToMapiMessageItem(), MapiNote)
            note2.Subject = "Pink color note"
            note2.Body = "This is a pink color note"
            note2.Color = NoteColor.Pink

            ' Note #3
            Dim note3 As MapiNote = DirectCast(message.ToMapiMessageItem(), MapiNote)
            note2.Subject = "Blue color note"
            note2.Body = "This is a blue color note"
            note2.Color = NoteColor.Blue
            note3.Height = 500
            note3.Width = 500

            Dim path As String = dataDir + "SampleNote_out.pst"

            If File.Exists(path) Then
                File.Delete(path)
            End If

            Using pst As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("SampleNote_out.pst"), FileFormatVersion.Unicode)
                Dim notesFolder As FolderInfo = pst.CreatePredefinedFolder("Notes", StandardIpmFolder.Notes)
                notesFolder.AddMapiMessageItem(note1)
                notesFolder.AddMapiMessageItem(note2)
                notesFolder.AddMapiMessageItem(note3)
            End Using
            ' ExEnd:CreateNewMapiCalendarAndAddToCalendarSubfolder
        End Sub
    End Class
End Namespace