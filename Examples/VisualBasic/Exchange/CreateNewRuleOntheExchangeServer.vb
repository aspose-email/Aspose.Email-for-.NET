Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports Aspose.Email.Mail.Calendaring

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class CreateNewRuleOntheExchangeServer
		Public Shared Sub Run()
			' ExStart:CreateNewRuleOntheExchangeServers
			' Set Exchange Server 2010 web service URL, Username, password, domain information
			Dim mailboxURI As String = "https://ex2010/ews/exchange.asmx"
			Dim username As String = "test.exchange"
			Dim password As String = "pwd"
			Dim domain As String = "ex2010.local"

			' Connect to the Exchange Server
			Dim credential As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxURI, credential)

			Console.WriteLine("Connected to Exchange server")

			Dim rule As New InboxRule()
			rule.DisplayName = "Message from client ABC"

			' Add conditions
			Dim newRules As New RulePredicates()
			' Set Subject contains string "ABC" and Add the conditions
			newRules.ContainsSubjectStrings.Add("ABC")
			newRules.FromAddresses.Add(New MailAddress("administrator@ex2010.local", True))
			rule.Conditions = newRules

			' Add Actions and Move the message to a folder
			Dim newActions As New RuleActions()
			newActions.MoveToFolder = "120:AAMkADFjMjNjMmNjLWE3NzgtNGIzNC05OGIyLTAwNTgzNjRhN2EzNgAuAAAAAABbwP+Tkhs0TKx1GMf0D/cPAQD2lptUqri0QqRtJVHwOKJDAAACL5KNAAA=AQAAAA=="
			rule.Actions = newActions
			client.CreateInboxRule(rule)
			' ExEnd:CreateNewRuleOntheExchangeServer
		End Sub
	End Class
End Namespace