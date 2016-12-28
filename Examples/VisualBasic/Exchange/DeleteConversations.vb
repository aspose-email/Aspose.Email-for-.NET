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
	Class DeleteConversations
		Public Shared Sub Run()
			Const  mailboxUri As String = "https://exchange/ews/exchange.asmx"
			Const  domain As String = ""
			Const  username As String = "username@ASE305.onmicrosoft.com"
			Const  password As String = "password"
			Dim credentials As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credentials)
			Console.WriteLine("Connected to Exchange 2010")
			' ExStart:DeleteConversations
			' Find those Conversation Items in the Inbox folder which we want to delete
			Dim conversations As ExchangeConversation() = client.FindConversations(client.MailboxInfo.InboxUri)
			For Each conversation As ExchangeConversation In conversations
				Console.WriteLine("Topic: " + conversation.ConversationTopic)
				' Delete the conversation item based on some condition
				If conversation.ConversationTopic.Contains("test email") = True Then
					client.DeleteConversationItems(conversation.ConversationId)
					Console.WriteLine("Deleted the conversation item")
				End If
			Next
			' ExEnd:DeleteConversations
		End Sub
	End Class
End Namespace