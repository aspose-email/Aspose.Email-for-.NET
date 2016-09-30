Imports Aspose.Email.Mail
Imports Aspose.Email.Mail.Bounce

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class GetDeliveryStatusNotificationMessages
		Public Shared Sub Run()
			' ExStart:GetDeliveryStatusNotificationMessages
			Dim fileName As String = RunExamples.GetDataDir_Email() + "failed1.msg"
			Dim mail As MailMessage = MailMessage.Load(fileName)
			Dim result As BounceResult = mail.CheckBounced()
			Console.WriteLine(fileName)
            Console.WriteLine("IsBounced : " + result.IsBounced.ToString())
            Console.WriteLine("Action : " + result.Action.ToString())
            Console.WriteLine("Recipient : " + result.Recipient.ToString())
			Console.WriteLine()
            Console.WriteLine("Reason : " + result.Reason.ToString())
            Console.WriteLine("Status : " + result.Status.ToString())
            Console.WriteLine("OriginalMessage ToAddress 1: " + result.OriginalMessage.[To](0).Address.ToString())
			Console.WriteLine()
			' ExEnd:GetDeliveryStatusNotificationMessages
		End Sub
	End Class
End Namespace