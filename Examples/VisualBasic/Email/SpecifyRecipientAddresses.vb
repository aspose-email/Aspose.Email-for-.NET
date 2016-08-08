Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class SpecifyRecipientAddresses
		Public Shared Sub Run()
			' ExStart:SpecifyRecipientAddresses
			' Declare message as MailMessage instance
			Dim message As New MailMessage()

			' Specify the recipients mail addresses
			message.[To].Add("receiver1@receiver.com")
			message.[To].Add("receiver2@receiver.com")
			message.[To].Add("receiver3@receiver.com")
			message.[To].Add("receiver4@receiver.com")

			' Specifying CC and BCC Addresses
			message.CC.Add("CC1@receiver.com")
			message.CC.Add("CC2@receiver.com")
			message.Bcc.Add("Bcc1@receiver.com")
			message.Bcc.Add("Bcc2@receiver.com")
			' ExEnd:SpecifyRecipientAddresses
		End Sub
	End Class
End Namespace