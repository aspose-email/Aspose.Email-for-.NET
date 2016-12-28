Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports Aspose.Email.Mail.Calendaring

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class FilterMessagesFromExchangeMailbox
		Public Shared Sub Run()
			' ExStart:FilterMessagesFromExchangeMailbox
			Try
				' Connect to Exchange Server
				Const  mailboxUri As String = "http://exchange-server/Exchange/username"
				Const  username As String = "username"
				Const  password As String = "password"
				Const  domain As String = "domain.com"

				Dim credential As New NetworkCredential(username, password, domain)
				Dim client As New ExchangeClient(mailboxUri, credential)

				' Query building by means of ExchangeQueryBuilder class
				Dim builder As New ExchangeQueryBuilder()

				' Set Subject and Emails that arrived today
				builder.Subject.Contains("Newsletter")
				builder.InternalDate.[On](DateTime.Now)

				Dim query As MailQuery = builder.GetQuery()
				' Get list of messages
				Dim messages As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri, query, False)
				Console.WriteLine("Imap: " + messages.Count + " message(s) found.")

				' Disconnect from IMAP
				client.Dispose()
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			End Try
			' ExEnd:FilterMessagesFromExchangeMailbox
		End Sub
	End Class
End Namespace