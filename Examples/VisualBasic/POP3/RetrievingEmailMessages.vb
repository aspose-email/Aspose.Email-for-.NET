Imports Aspose.Email.Mail
Imports Aspose.Email.Pop3

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class RetrievingEmailMessages
		Public Shared Sub Run()
			'ExStart:RetrievingEmailMessages
			' Create an instance of the Pop3Client class
			Dim client As New Pop3Client()

			' Specify host, username, password, Port and SecurityOptions for your client
			client.Host = "pop.gmail.com"
			client.Username = "your.username@gmail.com"
			client.Password = "your.password"
			client.Port = 995
			client.SecurityOptions = SecurityOptions.Auto
			Try
				Dim messageCount As Integer = client.GetMessageCount()
				' Create an instance of the MailMessage class and Retrieve message
				Dim message As MailMessage
				For i As Integer = 1 To messageCount
					message = client.FetchMessage(i)
                    Console.WriteLine("From:" + message.From.ToString())
					Console.WriteLine("Subject:" + message.Subject)
					Console.WriteLine(message.HtmlBody)
				Next
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			Finally
				client.Dispose()
			End Try
			'ExEnd:RetrievingEmailMessages
			Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ")
		End Sub
	End Class
End Namespace