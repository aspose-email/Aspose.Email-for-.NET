Imports Aspose.Email.Imap

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class ListingMessagesRecursively
		Public Shared Sub Run()
			' ExStart:ListingMessagesRecursively
			' Create an imapclient with host, user and password
			Dim client As New ImapClient()
			client.Host = "domain.com"
			client.Username = "username"
			client.Password = "password"
			client.SelectFolder("InBox")
			Dim msgsColl As ImapMessageInfoCollection = client.ListMessages(True)
			Console.WriteLine("Total Messages: " + msgsColl.Count)
			' ExEnd:ListingMessagesRecursively
		End Sub
	End Class
End Namespace