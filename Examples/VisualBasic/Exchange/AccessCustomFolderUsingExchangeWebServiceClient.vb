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
	Class AccessCustomFolderUsingExchangeWebServiceClient
		Public Shared Sub Run()
			' ExStart:AccessCustomFolderUsingExchangeWebServiceClient
			' Create instance of EWSClient class by giving credentials
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

			' Create ExchangeMailboxInfo, ExchangeMessageInfoCollection instance
			Dim mailbox As ExchangeMailboxInfo = client.GetMailboxInfo()
			Dim messages As ExchangeMessageInfoCollection = Nothing
			Dim subfolderInfo As New ExchangeFolderInfo()

			' Check if specified custom folder exisits and Get all the messages info from the target Uri
			client.FolderExists(mailbox.InboxUri, "TestInbox", subfolderInfo)
			messages = client.FindMessages(subfolderInfo.Uri)

			' Parse all the messages info collection
			For Each info As ExchangeMessageInfo In messages
				Dim strMessageURI As String = info.UniqueUri
				' now get the message details using FetchMessage()
				Dim msg As MailMessage = client.FetchMessage(strMessageURI)
				Console.WriteLine("Subject: " + msg.Subject)
			Next
			' ExEnd:AccessCustomFolderUsingExchangeWebServiceClient
		End Sub
	End Class
End Namespace