Imports System.Text
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddAttachmentsToMapiTask
        Public Shared Sub Run()
            ' ExStart:AddAttachmentsToMapiTask
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim attachmentContent As String = "Test attachment body"
            Dim attachmentName As String = "Test attachment name"

            Dim testTask As New MapiTask("Task with attacment", "Test body of task with attacment", DateTime.Now, DateTime.Now.AddHours(1))
            testTask.Attachments.Add(attachmentName, Encoding.Unicode.GetBytes(attachmentContent))
            testTask.Save(dataDir & Convert.ToString("AddAttachmentsToMapiTask_out"), TaskSaveFormat.Msg)
            ' ExEnd:AddAttachmentsToMapiTask
        End Sub
    End Class
End Namespace
