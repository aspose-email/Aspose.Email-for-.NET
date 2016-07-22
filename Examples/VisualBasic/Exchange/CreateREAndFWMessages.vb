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
	Class CreateREAndFWMessages
		Public Shared Sub Run()
			' ExStart:CreateREAndFWMessages
			Dim dataDir As String = RunExamples.GetDataDir_Exchange()

			Const  mailboxUri As String = "https://exchange.domain.com/ews/Exchange.asmx"
			Const  domain As String = ""
			Const  username As String = "username"
			Const  password As String = "password"
			Dim credential As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credential)

			Try
				Dim message As New MailMessage("user@domain.com", "user@domain.com", "TestMailRefw - " + Guid.NewGuid().ToString(), "TestMailRefw Implement ability to create RE and FW messages from source MSG file")
				client.Send(message)

				Dim messageInfoCol As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri)
				If messageInfoCol.Count = 1 Then
					Console.WriteLine("1 message in Inbox")
				Else
					Console.WriteLine("Error! No message in Inbox")
				End If

				Dim message1 As New MailMessage("user@domain.com", "user@domain.com", "TestMailRefw - " + Guid.NewGuid().ToString(), "TestMailRefw Implement ability to create RE and FW messages from source MSG file")

				client.Send(message1)
				messageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri)

				If messageInfoCol.Count = 2 Then
					Console.WriteLine("2 messages in Inbox")
				Else
					Console.WriteLine("Error! No new message in Inbox")
				End If

				Dim message2 As New MailMessage("user@domain.com", "user@domain.com", "TestMailRefw - " + Guid.NewGuid().ToString(), "TestMailRefw Implement ability to create RE and FW messages from source MSG file")
				message2.Attachments.Add(Attachment.CreateAttachmentFromString("Test attachment 1", "Attachment Name 1"))
				message2.Attachments.Add(Attachment.CreateAttachmentFromString("Test attachment 2", "Attachment Name 2"))

				' Reply, Replay All and forward Message
				client.Reply(message2, messageInfoCol(0))
				client.ReplyAll(message2, messageInfoCol(0))
				client.Forward(message2, messageInfoCol(0))
			Catch ex As Exception
				Console.WriteLine("Error in program" + ex.Message)
			End Try
			' ExEnd:CreateREAndFWMessages
		End Sub
	End Class
End Namespace