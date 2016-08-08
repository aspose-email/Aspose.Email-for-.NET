Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class RemoveAttachments
		Public Shared Sub Run()
			' ExStart:RemoveAttachments
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_Email()
			Dim dstEmailRemoved As String = dataDir & Convert.ToString("RemoveAttachments.msg")

			' Create an instance of MailMessage class
			Dim message As New MailMessage()
			message.From = "sender@sender.com"
			message.[To].Add("receiver@gmail.com")

			' Load an attachment
			Dim attachment As New Attachment(dataDir & Convert.ToString("1.txt"))

			' Add Multiple Attachment in instance of MailMessage class and Save message to disk
			message.Attachments.Add(attachment)
			message.AddAttachment(New Attachment(dataDir & Convert.ToString("1.jpg")))
			message.AddAttachment(New Attachment(dataDir & Convert.ToString("1.doc")))
			message.AddAttachment(New Attachment(dataDir & Convert.ToString("1.rar")))
			message.AddAttachment(New Attachment(dataDir & Convert.ToString("1.pdf")))

			' Remove attachment from your MailMessage and Save message to disk after removing a single attachment 
			message.Attachments.Remove(attachment)
			message.Save(dstEmailRemoved, SaveOptions.DefaultMsgUnicode)

			' Create a loop to display the no. of attachments present in email message
			For Each getAttachment As Attachment In message.Attachments
				' Save your attachments here and Display the the attachment file name
				getAttachment.Save((dataDir & Convert.ToString("/RemoveAttachments/")) + "attachment_out" + getAttachment.Name)
				Console.WriteLine(getAttachment.Name)
			Next
			' ExEnd:RemoveAttachments
			Console.WriteLine(Convert.ToString(Environment.NewLine + "Attachments removed successfully from ") & dstEmailRemoved)
		End Sub
	End Class
End Namespace