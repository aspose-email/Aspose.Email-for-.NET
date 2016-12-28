Imports System.Security.Cryptography.X509Certificates
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class EncryptAndDecryptMessage
		Public Shared Sub Run()
			' ExStart:EncryptAndDecryptMessage
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_Email()

			Dim publicCertFile As String = dataDir & Convert.ToString("MartinCertificate.cer")
			Dim privateCertFile As String = dataDir & Convert.ToString("MartinCertificate.pfx")

			Dim publicCert As New X509Certificate2(publicCertFile)
			Dim privateCert As New X509Certificate2(privateCertFile, "anothertestaccount")

			' Create a message
			Dim msg As New MailMessage("atneostthaecrcount@gmail.com", "atneostthaecrcount@gmail.com", "Test subject", "Test Body")

			' Encrypt the message
			Dim eMsg As MailMessage = msg.Encrypt(publicCert)
			If eMsg.IsEncrypted = True Then
				Console.WriteLine("Its encrypted")
			Else
				Console.WriteLine("Its NOT encrypted")
			End If

			' Decrypt the message
			Dim dMsg As MailMessage = eMsg.Decrypt(privateCert)
			If dMsg.IsEncrypted = True Then
				Console.WriteLine("Its encrypted")
			Else
				Console.WriteLine("Its NOT encrypted")
			End If
			' ExEnd:EncryptAndDecryptMessage
		End Sub
	End Class
End Namespace