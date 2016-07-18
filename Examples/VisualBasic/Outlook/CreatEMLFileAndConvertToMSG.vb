Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreatEMLFileAndConvertToMSG
        Public Shared Sub Run()
            ' ExStart:CreatEMLFileAndConvertToMSG
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create new message and save as EML
            Dim message As New MailMessage("from@domain.com", "to@domain.com")
            message.Subject = "subject of email"
            message.HtmlBody = "<b>Eml to msg conversion using Aspose.Email</b>" + "<br><hr><br><font color=blue>This is a test eml file which will be converted to msg format.</font>"
            ' Add attachments
            message.Attachments.Add(New Attachment(dataDir & Convert.ToString("attachment_1.doc")))
            message.Attachments.Add(New Attachment(dataDir & Convert.ToString("download.png")))

            ' Save it EML
            message.Save(dataDir & Convert.ToString("CreatEMLFileAndConvertToMSG_out.eml"), SaveOptions.DefaultEml)
            ' Save it to MSG
            message.Save(dataDir & Convert.ToString("CreatEMLFileAndConvertToMSG_out.msg"), SaveOptions.DefaultMsgUnicode)
            ' ExEnd:CreatEMLFileAndConvertToMSG
        End Sub
    End Class
End Namespace
