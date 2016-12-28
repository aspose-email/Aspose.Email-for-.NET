Imports Aspose.Email.Pop3

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
	Class GettingMailboxInfo
		Public Shared Sub Run()
			' Create an instance of the Pop3Client class
			'ExStart:GettingMailboxInfo
			Dim client As New Pop3Client()
			' Specify host, username, password, Port and SecurityOptions for your client
			client.Host = "pop.gmail.com"
			client.Username = "your.username@gmail.com"
			client.Password = "your.password"
			client.Port = 995
			client.SecurityOptions = SecurityOptions.Auto

			' Get the size of the mailbox,  Get mailbox info, number of messages in the mailbox
			Dim nSize As Long = client.GetMailboxSize()
			Console.WriteLine("Mailbox size is " + nSize + " bytes.")
			Dim info As Pop3MailboxInfo = client.GetMailboxInfo()
			Dim nMessageCount As Integer = info.MessageCount
			Console.WriteLine("Number of messages in mailbox are " + nMessageCount)
			Dim nOccupiedSize As Long = info.OccupiedSize
			Console.WriteLine("Occupied size is " + nOccupiedSize)
			'ExEnd:GettingMailboxInfo
			Console.WriteLine(Environment.NewLine + "Getting the mailbox information of POP3 server.")
		End Sub
	End Class
End Namespace