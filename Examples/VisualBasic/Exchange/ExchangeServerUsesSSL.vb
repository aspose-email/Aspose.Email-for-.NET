Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports Aspose.Email.Imap

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class ExchangeServerUsesSSL
		' ExStart:ExchangeServerUsesSSL
		Public Shared Sub Run()
			' Connect to Exchange Server using ImapClient class
			Dim imapClient As New ImapClient("ex07sp1", 993, "Administrator", "Evaluation1", New RemoteCertificateValidationCallback(AddressOf RemoteCertificateValidationHandler))
			imapClient.SecurityOptions = SecurityOptions.SSLExplicit

			' Select the Inbox folder
			imapClient.SelectFolder(ImapFolderInfo.InBox)

			' Get the list of messages
			Dim msgCollection As ImapMessageInfoCollection = imapClient.ListMessages()
			For Each msgInfo As ImapMessageInfo In msgCollection
				Console.WriteLine(msgInfo.Subject)
			Next
			' Disconnect from the server
			imapClient.Dispose()
		End Sub

		' Certificate verification handler
		Private Shared Function RemoteCertificateValidationHandler(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
			Return True
			' ignore the checks and go ahead
		End Function
		' ExEnd:ExchangeServerUsesSSL
	End Class
End Namespace