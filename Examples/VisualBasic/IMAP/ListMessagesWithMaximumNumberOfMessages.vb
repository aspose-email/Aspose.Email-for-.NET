Imports Aspose.Email.Imap

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class ListMessagesWithMaximumNumberOfMessages
		Public Shared Sub Run()
			' ExStart:ListMessagesWithMaximumNumberOfMessages
			' Create an imapclient with host, user and password
			Dim client As New ImapClient("localhost", "user", "password")

			' Select the inbox folder and Get the message info collection
			Dim builder As New ImapQueryBuilder()
			Dim query As MailQuery = builder.[Or](builder.[Or](builder.[Or](builder.[Or](builder.Subject.Contains(" (1) "), builder.Subject.Contains(" (2) ")), builder.Subject.Contains(" (3) ")), builder.Subject.Contains(" (4) ")), builder.Subject.Contains(" (5) "))
			Dim messageInfoCol4 As ImapMessageInfoCollection = client.ListMessages(query, 4)
			Console.WriteLine(If((messageInfoCol4.Count = 4), "Success", "Failure"))
			' ExEnd:ListMessagesWithMaximumNumberOfMessages
		End Sub
	End Class
End Namespace