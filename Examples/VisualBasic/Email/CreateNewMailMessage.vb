Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class CreateNewMailMessage
		Public Shared Sub Run()
			' ExStart:CreateNewMailMessage
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_Email()

			' Create a new instance of MailMessage class
			Dim message As New MailMessage()

			' Set subject of the message, Html body and sender information
			message.Subject = "New message created by Aspose.Email for .NET"
			message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/>" + "<font color=blue>This line is in blue color</font>"
			message.From = New MailAddress("from@domain.com", "Sender Name", False)

			' Add TO recipients and Add CC recipients
			message.[To].Add(New MailAddress("to1@domain.com", "Recipient 1", False))
			message.[To].Add(New MailAddress("to2@domain.com", "Recipient 2", False))
			message.CC.Add(New MailAddress("cc1@domain.com", "Recipient 3", False))
			message.CC.Add(New MailAddress("cc2@domain.com", "Recipient 4", False))

			' Save message in EML, EMLX, MSG and MHTML formats
			message.Save(dataDir & Convert.ToString("Message_out.eml"), SaveOptions.DefaultEml)
			message.Save(dataDir & Convert.ToString("Message_out.emlx"), SaveOptions.CreateSaveOptions(MailMessageSaveType.EmlxFormat))
			message.Save(dataDir & Convert.ToString("Message_out.msg"), SaveOptions.DefaultMsgUnicode)
			message.Save(dataDir & Convert.ToString("Message_out.mhtml"), SaveOptions.DefaultMhtml)
			' ExEnd:CreateNewMailMessage
		End Sub
	End Class
End Namespace