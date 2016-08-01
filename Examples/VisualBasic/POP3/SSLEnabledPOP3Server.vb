Imports Aspose.Email.Pop3

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class SSLEnabledPOP3Server
		Public Shared Sub Run()
			'ExStart:SSLEnabledPOP3Server
			' Create an instance of the Pop3Client class
			Dim client As New Pop3Client()
			' Specify host, username and password, Port and  SecurityOptions for your client
			client.Host = "pop.gmail.com"
			client.Username = "your.username@gmail.com"
			client.Password = "your.password"
			client.Port = 995
			client.SecurityOptions = SecurityOptions.Auto
			Console.WriteLine(Environment.NewLine + "Connecting to POP3 server using SSL.")
			'ExEnd:SSLEnabledPOP3Server
		End Sub
	End Class
End Namespace