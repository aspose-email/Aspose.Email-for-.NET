Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class MessagesFromIMAPServerToDisk
        Public Shared Sub Run()
            ' ExStart:MessagesFromIMAPServerToDisk
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_IMAP()

            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto
            Try
                ' Select the inbox folder and Get the message info collection
                client.SelectFolder(ImapFolderInfo.InBox)
                Dim list As ImapMessageInfoCollection = client.ListMessages()

                ' Download each message
                For i As Integer = 0 To list.Count - 1
                    ' Save the EML file locally
                    client.SaveMessage(list(i).UniqueId, dataDir + list(i).UniqueId & Convert.ToString(".eml"))
                Next
                ' Disconnect to the remote IMAP server
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            ' ExStart:MessagesFromIMAPServerToDisk
            Console.WriteLine(Environment.NewLine + "Downloaded messages from IMAP server.")
        End Sub
    End Class
End Namespace