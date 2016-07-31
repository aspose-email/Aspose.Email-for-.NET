Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class ConnectingWithIMAPServer
        Public Shared Sub Run()
            ' ExStart:ConnectingWithIMAPServer
            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto
            Try
                ' Disconnect to the remote IMAP server
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            Console.WriteLine(Environment.NewLine + "Connected to IMAP server.")
            ' ExEnd:ConnectingWithIMAPServer
        End Sub
    End Class
End Namespace