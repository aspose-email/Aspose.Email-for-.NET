Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddMapiNoteToPST
        Public Shared Sub Run()
            ' ExStart:AddMapiNoteToPST
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim checkfile As String = dataDir + "Note.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If

            Dim mess As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("Note.msg"))

            ' Create three Notes  
            Dim note1 As MapiNote = DirectCast(mess.ToMapiMessageItem(), MapiNote)
            note1.Subject = "Yellow color note"
            note1.Body = "This is a yellow color note"

            Dim note2 As MapiNote = DirectCast(mess.ToMapiMessageItem(), MapiNote)
            note2.Subject = "Pink color note"
            note2.Body = "This is a pink color note"
            note2.Color = NoteColor.Pink

            Dim note3 As MapiNote = DirectCast(mess.ToMapiMessageItem(), MapiNote)
            note2.Subject = "Blue color note"
            note2.Body = "This is a blue color note"
            note2.Color = NoteColor.Blue
            note3.Height = 500
            note3.Width = 500

            Dim checkfile1 As String = dataDir + "AddMapiNoteToPST_out.pst"

            If File.Exists(checkfile1) Then
                File.Delete(checkfile1)
            Else
            End If

            Using personalStorage1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("AddMapiNoteToPST_out.pst"), FileFormatVersion.Unicode)
                Dim notesFolder As FolderInfo = personalStorage1.CreatePredefinedFolder("Notes", StandardIpmFolder.Notes)
                notesFolder.AddMapiMessageItem(note1)
                notesFolder.AddMapiMessageItem(note2)
                notesFolder.AddMapiMessageItem(note3)
            End Using
            ' ExEnd:AddMapiNoteToPST
        End Sub
    End Class
End Namespace
