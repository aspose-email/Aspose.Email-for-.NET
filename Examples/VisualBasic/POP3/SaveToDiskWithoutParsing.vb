Imports Aspose.Email.Pop3

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class SaveToDiskWithoutParsing
		Public Shared Sub Run()
			'ExStart:SaveToDiskWithoutParsing
			' The path to the File directory.
			Dim dataDir As String = RunExamples.GetDataDir_POP3()
			Dim dstEmail As String = dataDir & Convert.ToString("InsertHeaders.eml")

			' Create an instance of the Pop3Client class
			Dim client As New Pop3Client()

			' Specify host, username, password, Port and SecurityOptions for your client
			client.Host = "pop.gmail.com"
			client.Username = "your.username@gmail.com"
			client.Password = "your.password"
			client.Port = 995
			client.SecurityOptions = SecurityOptions.Auto

			Try
				' Save message to disk by message sequence number
				client.SaveMessage(1, dstEmail)
				client.Dispose()
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			End Try
			Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ")
			'ExEnd:SaveToDiskWithoutParsing
		End Sub
	End Class
End Namespace