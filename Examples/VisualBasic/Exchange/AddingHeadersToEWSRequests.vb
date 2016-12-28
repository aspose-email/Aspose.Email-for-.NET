Imports System.Collections.Generic
Imports System.Configuration
Imports System.Net
Imports System.Threading
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class AddingHeadersToEWSRequests
		Public Shared Sub Run()
			' ExStart:AddingHeadersToEWSRequests
			Using client As IEWSClient = EWSClient.GetEWSClient("exchange.domain.com/ews/Exchange.asmx", "username", "password", "")
				client.AddHeader("X-AnchorMailbox", "username@domain.com")
				Dim messageInfoCol As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri)
			End Using
			' ExEnd:AddingHeadersToEWSRequests
		End Sub
	End Class
End Namespace