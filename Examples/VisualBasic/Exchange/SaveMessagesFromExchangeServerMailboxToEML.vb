Imports Aspose.Email.Exchange

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class SaveMessagesFromExchangeServerMailboxToEML
		Public Shared Sub Run()
			' ExStart:EnumeratMessagesWithPaginginEWS
			Dim dataDir As String = RunExamples.GetDataDir_Exchange()
			' Create instance of ExchangeClient class by giving credentials
			Dim client As New ExchangeClient("http://Servername/exchange/username", "username", "password", "domain")

			' Call ListMessages method to list messages info from Inbox
			Dim msgCollection As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri)

			' Loop through the collection to get Message URI
			For Each msgInfo As ExchangeMessageInfo In msgCollection
				Dim strMessageURI As String = msgInfo.UniqueUri

				' Now save the message in disk
				client.SaveMessage(strMessageURI, dataDir + msgInfo.MessageId & Convert.ToString("_out.eml"))
			Next
			' ExEnd:EnumeratMessagesWithPaginginEWS
		End Sub
	End Class
End Namespace