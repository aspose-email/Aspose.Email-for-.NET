Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class RenamingFolders
        Public Shared Sub Run()
            'ExStart:RenamingFolders
            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username and password, port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto
            Try
                ' Rename a folder and Disconnect to the remote IMAP server
                client.RenameFolder("Aspose", "Client")
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            'ExEnd:RenamingFolders
            Console.WriteLine(Environment.NewLine + "Renamed folders on IMAP server.")
        End Sub
    End Class
End Namespace