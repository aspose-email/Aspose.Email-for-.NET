Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class DeletingFolders
        Public Shared Sub Run()
            'ExStart:DeletingFolders
            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username and password, and set port for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto
            Try
                ' Rename a folder and Disconnect to the remote IMAP server
                client.DeleteFolder("Client")
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            'ExEnd:DeletingFolders
            Console.WriteLine(Environment.NewLine + "Deleted folders on IMAP server.")
        End Sub
    End Class
End Namespace