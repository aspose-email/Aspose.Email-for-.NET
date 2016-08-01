Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class SSLEnabledIMAPServer
        Public Shared Sub Run()
            'ExStart:SSLEnabledIMAPServer
            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()
            ' Specify host, username and password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto

            Try
                Console.WriteLine("Logged in to the IMAP server")
                ' Disconnect to the remote IMAP server
                client.Dispose()
                Console.WriteLine("Disconnected from the IMAP server")
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            'ExEnd:SSLEnabledIMAPServer
            Console.WriteLine(Environment.NewLine + "Connected to IMAP server with SSL.")
        End Sub
    End Class
End Namespace