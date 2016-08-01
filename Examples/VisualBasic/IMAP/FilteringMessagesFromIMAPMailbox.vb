Imports Aspose.Email.Imap

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class FilteringMessagesFromIMAPMailbox
		Public Shared Sub Run()
			' ExStart:FilteringMessagesFromIMAPMailbox
			' Connect and log in to IMAP
			Const  host As String = "host"
			Const  port As Integer = 143
			Const  username As String = "user@host.com"
			Const  password As String = "password"
			Dim client As New ImapClient(host, port, username, password)
			client.SelectFolder("Inbox")
			' Set conditions, Subject contains "Newsletter", Emails that arrived today
			Dim builder As New ImapQueryBuilder()
			builder.Subject.Contains("Newsletter")
			builder.InternalDate.[On](DateTime.Now)
			' Build the query and Get list of messages
			Dim query As MailQuery = builder.GetQuery()
			Dim messages As ImapMessageInfoCollection = client.ListMessages(query)
			Console.WriteLine("Imap: " + messages.Count + " message(s) found.")
			' Disconnect from IMAP
			client.Dispose()
			' ExEnd:FilteringMessagesFromIMAPMailbox
		End Sub
	End Class
End Namespace