Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class AddingNewMessage
        Public Shared Sub Run()
            'ExStart:AddingNewMessage
            ' Create a message
            Dim msg As New MailMessage("user@domain1.com", "user@domain2.com", "subject", "message")

            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username, password, port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto
            Try
                ' Subscribe to the Inbox folder, Append the newly created message and Disconnect to the remote IMAP server
                client.SelectFolder(ImapFolderInfo.InBox)
                client.SubscribeFolder(client.CurrentFolder.Name)
                client.AppendMessage(client.CurrentFolder.Name, msg)
                Console.WriteLine("New Message Added Successfully")
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            Console.WriteLine(Environment.NewLine + "Added new message on IMAP server.")
            'ExEnd:AddingNewMessage
        End Sub
    End Class
End Namespace