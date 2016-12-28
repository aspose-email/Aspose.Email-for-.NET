Imports Aspose.Email.Imap

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class GettingFoldersInformation
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
                ' Get all folders in the currently subscribed folder
                Dim folderInfoColl As ImapFolderInfoCollection = client.ListFolders()

                ' Iterate through the collection to get folder info one by one
                For Each folderInfo As ImapFolderInfo In folderInfoColl
                    ' Folder name and get New messages in the folder
                    Console.WriteLine("Folder name is " + folderInfo.Name)
                    Dim folderExtInfo As ImapFolderInfo = client.GetFolderInfo(folderInfo.Name)
                    Console.WriteLine("New message count: " + folderExtInfo.NewMessageCount)
                    Console.WriteLine("Is it readonly? " + folderExtInfo.[ReadOnly])
                    Console.WriteLine("Total number of messages " + folderExtInfo.TotalMessageCount)
                Next
                ' Disconnect to the remote IMAP server
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try
            ' ExEnd:ConnectingWithIMAPServer
            Console.WriteLine(Environment.NewLine + "Getting folders information from IMAP server.")
        End Sub
    End Class
End Namespace