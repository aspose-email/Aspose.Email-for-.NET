Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class DeleteMessagesFromExchangeServer
		Public Shared Sub Run()
			' ExStart:DeleteMessagesFromExchangeServer
			' Create instance of IEWSClient class by giving credentials
			Dim mailboxURI As String = "https://Ex2003/exchange/administrator"
			' WebDAV
			Dim username As String = "administrator"
			Dim password As String = "pwd"
			Dim domain As String = "domain.local"

			Console.WriteLine("Connecting to Exchange Server....")
			Dim credential As New NetworkCredential(username, password, domain)
			Dim client As New ExchangeClient(mailboxURI, credential)

			Dim mailboxInfo As ExchangeMailboxInfo = client.GetMailboxInfo()

			' List all messages from Inbox folder
			Console.WriteLine("Listing all messages from Inbox....")
			Dim msgInfoColl As ExchangeMessageInfoCollection = client.ListMessages(mailboxInfo.InboxUri)
			For Each msgInfo As ExchangeMessageInfo In msgInfoColl
				' Delete message based on some criteria
				If msgInfo.Subject IsNot Nothing AndAlso msgInfo.Subject.ToLower().Contains("delete") = True Then
					client.DeleteMessage(msgInfo.UniqueUri)
					Console.WriteLine("Message deleted...." + msgInfo.Subject)
						' Do something else
				Else
				End If
			Next
			' ExEnd:DeleteMessagesFromExchangeServer
		End Sub
	End Class
End Namespace