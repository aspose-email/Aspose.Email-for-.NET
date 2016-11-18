Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ConvertMSGToMIMEMessage
        Public Shared Sub Run()
            'ExStart:ConvertMSGToMIMEMessage
            Dim msg As New MapiMessage("sender@test.com", "recipient1@test.com; recipient2@test.com", "Test Subject", "This is a body of message.")
            Dim options = New MailConversionOptions()
            options.ConvertAsTnef = True
            Dim mail As MailMessage = msg.ToMailMessage(options)
            'ExEnd:ConvertMSGToMIMEMessage
        End Sub
    End Class
End Namespace