Imports System.Diagnostics
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
	Class SendEmailUsingSMTP
		' ExStart:SendEmailUsingSMTP
		Public Shared Sub Run()
			' Declare msg as MailMessage instance
			Dim msg As New MailMessage()

			' Create an instance of SmtpClient class
			Dim client As New SmtpClient()

			' Specify your mailing host server, Username, Password, Port # and Security option
			client.Host = "mail.server.com"
			client.Username = "username"
			client.Password = "password"
			client.Port = 587
			client.SecurityOptions = SecurityOptions.SSLExplicit
			Try
				' Client.Send will send this message
				client.Send(msg)
				Console.WriteLine("Message sent")
			Catch ex As Exception
				Trace.WriteLine(ex.ToString())
			End Try
			Console.WriteLine("Press enter to quit")
			Console.Read()
		End Sub
		' ExEnd:SendEmailUsingSMTP
	End Class
End Namespace