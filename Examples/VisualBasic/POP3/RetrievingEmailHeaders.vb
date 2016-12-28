Imports Aspose.Email.Mime
Imports Aspose.Email.Pop3

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class RetrievingEmailHeaders
		Public Shared Sub Run()
			' ExStart:RetrievingEmailHeaders
			' Create an instance of the Pop3Client class
			Dim client As New Pop3Client()

			' Specify host, username. password, Port and SecurityOptions for your client
			client.Host = "pop.gmail.com"
			client.Username = "your.username@gmail.com"
			client.Password = "your.password"
			client.Port = 995
			client.SecurityOptions = SecurityOptions.Auto
			Try
				Dim headers As HeaderCollection = client.GetMessageHeaders(1)
				For i As Integer = 0 To headers.Count - 1
					' Display key and value in the header collection
					Console.Write(headers.Keys(i))
					Console.Write(" : ")
					Console.WriteLine(headers.[Get](i))
				Next
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			Finally
				client.Dispose()
			End Try
			' ExEnd:RetrievingEmailHeaders
			Console.WriteLine(Environment.NewLine + "Displayed header information from emails using POP3 ")
		End Sub
	End Class
End Namespace