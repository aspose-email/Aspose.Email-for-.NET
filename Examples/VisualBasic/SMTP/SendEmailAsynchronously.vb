Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class SendEmailAsynchronously
		' ExStart:SendEmailAsynchronously
		Public Shared Sub Run()
			SendMail()
		End Sub

		Private Shared Function GetSmtpClient2() As SmtpClient
			Dim client As New SmtpClient()
			client.Host = "smtp.gmail.com"
			'Specify your mail Username, Password, Port # and security option
			client.Username = "user"
			client.Password = "password"
			client.Port = 587
			client.SecurityOptions = SecurityOptions.SSLExplicit
			Return client
		End Function
		Private Shared Sub SendMail()
			'Declare msg as MailMessage instance
			Dim msg As New MailMessage("sender@gmail.com", "receiver@gmail.com", "Test subject", "Test body")
			Dim client As SmtpClient = GetSmtpClient2()
			Dim state As New Object()
			Dim ar As IAsyncResult = client.BeginSend(msg, Callback, state)

			Console.WriteLine("Sending message... press c to cancel mail. Press any other key to exit.")
			Dim answer As String = Console.ReadLine()

			' If the user canceled the send, and mail hasn't been sent yet,
			If answer IsNot Nothing AndAlso answer.StartsWith("c") Then
				client.CancelAsyncOperation(ar)
			End If

			msg.Dispose()
			Console.WriteLine("Goodbye.")
		End Sub
		Shared Callback As AsyncCallback = Sub(ar As IAsyncResult) 
		Dim task = TryCast(ar, IAsyncResultExt)
		If task.IsCanceled Then
			Console.WriteLine("Send canceled.")
		End If

		If task.ErrorInfo IsNot Nothing Then
			Console.WriteLine("{0}", task.ErrorInfo)
		Else
			Console.WriteLine("Message Sent.")
		End If

End Sub
		' ExEnd:SendEmailAsynchronously
	End Class
End Namespace