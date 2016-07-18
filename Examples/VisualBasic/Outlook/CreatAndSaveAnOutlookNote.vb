Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreatAndSaveAnOutlookNote
        Public Shared Sub Run()
            ' ExStart:CreatAndSaveAnOutlookNote
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            ' Create MapiNote and set Properties 
            Dim note3 As New MapiNote()
            note3.Subject = "Blue color note"
            note3.Body = "This is a blue color note"
            note3.Color = NoteColor.Blue
            note3.Height = 500
            note3.Width = 500
            note3.Save(dataDir & Convert.ToString("MapiNote_out.msg"), NoteSaveFormat.Msg)
            ' ExEnd:CreatAndSaveAnOutlookNote
        End Sub
    End Class
End Namespace
