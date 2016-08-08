Imports System.Security.Cryptography.X509Certificates
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class SignAMessage
        Public Shared Sub Run()
            ' ExStart:SignAMessage
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim publicCertFile As String = dataDir & Convert.ToString("MartinCertificate.cer")
            Dim privateCertFile As String = dataDir & Convert.ToString("MartinCertificate.pfx")
            Dim publicCert As New X509Certificate2(publicCertFile)
            Dim privateCert As New X509Certificate2(privateCertFile, "password")
            Dim msg As New MailMessage("userfrom@gmail.com", "userto@gmail.com", "Signed message only", "Test Body of signed message")
            Dim signed As MailMessage = msg.AttachSignature(privateCert)
            Dim encrypted As MailMessage = signed.Encrypt(publicCert)
            Dim decrypted As MailMessage = encrypted.Decrypt(privateCert)
            Dim unsigned As MailMessage = decrypted.RemoveSignature()
            'The original message with proper body
            Dim mapi As MapiMessage = MapiMessage.FromMailMessage(unsigned)
            ' ExEnd:SignAMessage
        End Sub
    End Class
End Namespace