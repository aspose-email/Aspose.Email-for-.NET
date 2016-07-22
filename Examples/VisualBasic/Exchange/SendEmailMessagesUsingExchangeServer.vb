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
	Class SendEmailMessagesUsingExchangeServer
		Public Shared Sub Run()
			' ExStart:SendEmailMessagesUsingExchangeServer
			' Create instance of ExchangeClient class by giving credentials
			Dim client As New ExchangeClient("https://MachineName/exchange/username", "username", "password", "domain")

			' Create instance of type MailMessage
			Dim msg As New MailMessage()
			msg.From = "sender@domain.com"
			msg.[To] = "recipient@ domain.com "
			msg.Subject = "Sending message from exchange server"
			msg.HtmlBody = "<h3>sending message from exchange server</h3>"

			' Send the message
			client.Send(msg)
			' ExEnd:SendEmailMessagesUsingExchangeServer
		End Sub
	End Class
End Namespace