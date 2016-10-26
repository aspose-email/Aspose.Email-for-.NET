Imports System.Text
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class SetDefaultTextEncoding
        Public Shared Sub Run()
            ' ExStart:SetDefaultTextEncoding           
            ' Create an instance of MailMessage
            Dim fileName As String = RunExamples.GetDataDir_Email()
            Dim msg As New MailMessage()

            ' Set the default or preferred encoding. This encoding will be used as the default for the from/to email addresses, subject and body of message.
            msg.PreferredTextEncoding = Encoding.GetEncoding(28591)

            ' Set email addresses, subject and body
            msg.From = New MailAddress("dmo@domain.com", "démo")
            msg.[To].Add(New MailAddress("dmo@domain.com", "démo"))
            msg.Subject = "démo"
            msg.HtmlBody = "démo"
            msg.Save(fileName & Convert.ToString("SetDefaultTextEncoding_out.msg"), SaveOptions.DefaultMsg)
            ' ExEnd:SetDefaultTextEncoding
        End Sub
    End Class
End Namespace