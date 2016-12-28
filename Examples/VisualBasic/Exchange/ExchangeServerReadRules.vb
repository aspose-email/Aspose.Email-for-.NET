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
	Class ExchangeServerReadRules
		Public Shared Sub Run()
			' ExStart:ExchangeServerReadRules
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

			' Display information about each rule
			For Each inboxRule As InboxRule In inboxRules
				Console.WriteLine("Display Name: " + inboxRule.DisplayName)

				' Check if there is a "From Address" condition
				If inboxRule.Conditions.FromAddresses.Count > 0 Then
					For Each fromAddress As MailAddress In inboxRule.Conditions.FromAddresses
						Console.WriteLine("From: " + fromAddress.DisplayName + " - " + fromAddress.Address)
					Next
				End If
				' Check if there is a "Subject Contains" condition
				If inboxRule.Conditions.ContainsSubjectStrings.Count > 0 Then
					For Each subject As [String] In inboxRule.Conditions.ContainsSubjectStrings
						Console.WriteLine("Subject contains: " + subject)
					Next
				End If
				' Check if there is a "Move to Folder" action
				If inboxRule.Actions.MoveToFolder.Length > 0 Then
					Console.WriteLine("Move message to folder: " + inboxRule.Actions.MoveToFolder)
				End If
			Next
			' ExEnd:ExchangeServerReadRules
		End Sub
	End Class
End Namespace
