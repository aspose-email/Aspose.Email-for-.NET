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
	Class UpdateRuleOntheExchangeServer
		Public Shared Sub Run()
			' ExStart:UpdateRuleOntheExchangeServer           
			' Set mailboxURI, Username, password, domain information
			Dim mailboxURI As String = "https://ex2010/ews/exchange.asmx"
			Dim username As String = "test.exchange"
			Dim password As String = "pwd"
			Dim domain As String = "ex2010.local"

			' Connect to the Exchange Server
			Dim credential As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxURI, credential)

			Console.WriteLine("Connected to Exchange server")

			' Get all Inbox Rules
			Dim inboxRules As InboxRule() = client.GetInboxRules()

			' Loop through each rule
			For Each inboxRule As InboxRule In inboxRules
				Console.WriteLine("Display Name: " + inboxRule.DisplayName)
				If inboxRule.DisplayName = "Message from client ABC" Then
					Console.WriteLine("Updating the rule....")
					inboxRule.Conditions.FromAddresses(0) = New MailAddress("administrator@ex2010.local", True)
					client.UpdateInboxRule(inboxRule)
				End If
			Next
			' ExEnd:UpdateRuleOntheExchangeServer
		End Sub
	End Class
End Namespace