Imports Aspose.Email.Pop3

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class FilterMessagesFromPOP3Mailbox
		Public Shared Sub Run()
			' ExStart:FilterMessagesFromPOP3Mailbox
			' Connect and log in to POP3
			Const  host As String = "host"
			Const  port As Integer = 110
			Const  username As String = "user@host.com"
			Const  password As String = "password"
			Dim client As New Pop3Client(host, port, username, password)

			' Set conditions, Subject contains "Newsletter", Emails that arrived today
			Dim builder As New MailQueryBuilder()
			builder.Subject.Contains("Newsletter")
			builder.InternalDate.[On](DateTime.Now)
			' Build the query and Get list of messages
			Dim query As MailQuery = builder.GetQuery()
			Dim messages As Pop3MessageInfoCollection = client.ListMessages(query)
			Console.WriteLine("Pop3: " + messages.Count + " message(s) found.")
			' ExEnd:FilterMessagesFromPOP3Mailbox
		End Sub
	End Class
End Namespace