Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ConvertMIMEMessageToEML
        Public Shared Sub Run()
            ' ExStart:ConvertMIMEMessageToEML
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim msg As MailMessage = MailMessage.Load(dataDir & Convert.ToString("Message2.msg"))
            ' Save File to disk
            msg.Save(dataDir & Convert.ToString("ConvertMIMEMessageToEML_out.eml"))
            ' ExEnd:ConvertMIMEMessageToEML
        End Sub
    End Class
End Namespace