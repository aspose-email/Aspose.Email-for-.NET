Imports System.Net
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
	Class CreateAndSendingMessageWithVotingOptions
		Public Shared Sub Run()
			' ExStart:CreateAndSendingMessageWithVotingOptions
			Dim address As String = "firstname.lastname@aspose.com"

			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
			Dim message As MailMessage = CreateTestMessage(address)

			' Set FollowUpOptions Buttons
			Dim options As New FollowUpOptions()
			options.VotingButtons = "YesNoMaybeExactly!"

			client.Send(message, options)
			' ExEnd:CreateAndSendingMessageWithVotingOptions
		End Sub

		' ExStart:CreateTestMessage
		Private Shared Function CreateTestMessage(address As String) As MailMessage
			Dim eml As New MailMessage(address, address, "Flagged message", "Make it nice and short, but descriptive. The description may appear in search engines' search results pages...")

			Return eml
		End Function
		' ExEnd:CreateTestMessage
	End Class
End Namespace