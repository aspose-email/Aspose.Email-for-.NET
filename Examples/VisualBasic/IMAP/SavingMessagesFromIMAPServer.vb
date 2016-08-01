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
	Class SavingMessagesFromIMAPServer
		Public Shared Sub Run()
			' ExStart:SavingMessagesFromIMAPServer
			' The path to the file directory.
			Dim dataDir As String = RunExamples.GetDataDir_IMAP()

			' Create an imapclient with host, user and password
			Dim client As New ImapClient("localhost", "user", "password")

			' Select the inbox folder and Get the message info collection
			client.SelectFolder(ImapFolderInfo.InBox)
			Dim list As ImapMessageInfoCollection = client.ListMessages()

			' Download each message
			For i As Integer = 0 To list.Count - 1
				' Save the message in MSG format
				Dim message As MailMessage = client.FetchMessage(list(i).UniqueId)
				message.Save(dataDir + list(i).UniqueId & Convert.ToString("_out.msg"), SaveOptions.DefaultMsgUnicode)
			Next
			' ExEnd:SavingMessagesFromIMAPServer
		End Sub
	End Class
End Namespace