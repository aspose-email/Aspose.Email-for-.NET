Imports System.Threading
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class RetrievingMessagesAsynchronously
		Public Shared Sub Run()
			' ExStart:RetrievingMessagesAsynchronously
			' Connect and log in to IMAP
			Using client As New ImapClient("host", "username", "password")
				client.SelectFolder("Issues/SubFolder")
				Dim messages As ImapMessageInfoCollection = client.ListMessages()
				Dim evnt As New AutoResetEvent(False)
				Dim message As MailMessage = Nothing
				Dim callback As AsyncCallback = Sub(ar As IAsyncResult) 
				message = client.EndFetchMessage(ar)
				evnt.[Set]()
			End Sub
				client.BeginFetchMessage(messages(0).SequenceNumber, callback, Nothing)
				evnt.WaitOne()
			End Using
			' ExEnd:RetrievingMessagesAsynchronously
		End Sub
	End Class
End Namespace
