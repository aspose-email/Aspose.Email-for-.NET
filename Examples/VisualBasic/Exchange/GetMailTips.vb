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
	Class GetMailTips
		Public Shared Sub Run()
			' ExStart:GetMailTips
			' Create instance of EWSClient class by giving credentials
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
			Console.WriteLine("Connected to Exchange server...")
			' Provide mail tips options
			Dim addrColl As New MailAddressCollection()
			addrColl.Add(New MailAddress("test.exchange@ex2010.local", True))
			addrColl.Add(New MailAddress("invalid.recipient@ex2010.local", True))
			Dim options As New GetMailTipsOptions("administrator@ex2010.local", addrColl, MailTipsType.All)

			' Get Mail Tips
			Dim tips As MailTips() = client.GetMailTips(options)

			' Display information about each Mail Tip
			For Each tip As MailTips In tips
				' Display Out of office message, if present
				If tip.OutOfOffice IsNot Nothing Then
					Console.WriteLine("Out of office: " + tip.OutOfOffice.ReplyBody.Message)
				End If

				' Display the invalid email address in recipient, if present
				If tip.InvalidRecipient = True Then
                    Console.WriteLine("The recipient address is invalid: " + tip.RecipientAddress.ToString())
				End If
			Next
			' ExEnd:GetMailTips
		End Sub
	End Class
End Namespace