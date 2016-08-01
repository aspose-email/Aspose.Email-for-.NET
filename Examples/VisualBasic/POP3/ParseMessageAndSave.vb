Imports Aspose.Email.Mail
Imports Aspose.Email.Pop3

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class ParseMessageAndSave
		Public Shared Sub Run()
			' ExStart:ParseMessageAndSave
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_POP3()

			' Create an instance of the Pop3Client class
			Dim client As New Pop3Client()

			' Specify host, username and password, Port and SecurityOptions for your client
			client.Host = "pop.gmail.com"
			client.Username = "your.username@gmail.com"
			client.Password = "your.password"
			client.Port = 995
			client.SecurityOptions = SecurityOptions.Auto

			Try
				' Fetch the message by its sequence number and Save the message using its subject as the file name
				Dim msg As MailMessage = client.FetchMessage(1)
				msg.Save(dataDir & Convert.ToString("first-message_out.eml"), SaveOptions.DefaultEml)
				client.Dispose()
			Catch ex As Exception
				Console.WriteLine(Environment.NewLine + ex.Message)
			Finally
				client.Dispose()
			End Try
			' ExEnd:ParseMessageAndSave
			Console.WriteLine((Convert.ToString(Environment.NewLine + "Downloaded email using POP3. Message saved at ") & dataDir) + "first-message_out.eml")
		End Sub
	End Class
End Namespace