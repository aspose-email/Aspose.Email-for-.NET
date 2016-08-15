Imports System.IO
Imports System.Security.Cryptography
Imports Aspose.Email.Mail
Imports Aspose.Email.Mail.DKIM

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'
Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class SignEmailsWithDKIM
        Public Shared Sub Run()
            Dim privateKeyFile As String = Path.Combine(RunExamples.GetDataDir_SMTP().Replace("_Send", String.Empty), RunExamples.GetDataDir_SMTP() + "key2.pem")

            Dim rsa As RSACryptoServiceProvider = PemReader.GetPrivateKey(privateKeyFile)
            Dim signInfo As New DKIMSignatureInfo("test", "yandex.ru")
            signInfo.Headers.Add("From")
            signInfo.Headers.Add("Subject")

            Dim mailMessage As New MailMessage("useremail@gmail.com", "test@gmail.com")
            mailMessage.Subject = "Signed DKIM message text body"
            mailMessage.Body = "This is a text body signed DKIM message"
            Dim signedMsg As MailMessage = mailMessage.DKIMSign(rsa, signInfo)

            Try
                Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
                client.Send(signedMsg)
            Finally
            End Try
        End Sub
    End Class
End Namespace