Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class SettingMessageFlags
        Public Shared Sub Run()
            'ExStart:SettingMessageFlags
            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto

            Try
                Console.WriteLine("Logged in to the IMAP server")
                ' Mark the message as read and Disconnect to the remote IMAP server
                client.ChangeMessageFlags(1, ImapMessageFlags.IsRead)
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            'ExEnd:SettingMessageFlags
            Console.WriteLine(Environment.NewLine + "Set message flags from IMAP server.")
        End Sub
    End Class
End Namespace