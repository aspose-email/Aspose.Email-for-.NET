Imports System.IO
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook
Imports Aspose.Email.Pop3
Imports Aspose.Email
Imports Aspose.Email.Mime
Imports Aspose.Email.Imap
Imports System.Configuration
Imports System.Data
Imports Aspose.Email.Mail.Bounce
Imports Aspose.Email.Exchange

Namespace Aspose.Email.Examples.VisualBasic.IMAP
    Public Class GettingFoldersInformation
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_IMAP()
            Dim dstEmail As String = dataDir & Convert.ToString("1234.eml")

            'Create an instance of the ImapClient class
            Dim client As New ImapClient()

            'Specify host, username and password for your client
            client.Host = "imap.gmail.com"

            ' Set username
            client.Username = "your.username@gmail.com"

            ' Set password
            client.Password = "your.password"

            ' Set the port to 993. This is the SSL port of IMAP server
            client.Port = 993

            ' Enable SSL
            client.SecurityOptions = SecurityOptions.Auto

            Try
                ' Get all folders in the currently subscribed folder
                Dim folderInfoColl As Aspose.Email.Imap.ImapFolderInfoCollection = client.ListFolders()

                ' Iterate through the collection to get folder info one by one
                For Each folderInfo As Aspose.Email.Imap.ImapFolderInfo In folderInfoColl
                    ' Folder name
                    Console.WriteLine("Folder name is " + folderInfo.Name)
                    Dim folderExtInfo As ImapFolderInfo = client.ListFolder(folderInfo.Name)
                    ' New messages in the folder
                    Console.WriteLine("New message count: " + folderExtInfo.NewMessageCount)
                    ' Check whether its readonly
                    Console.WriteLine("Is it readonly? " + folderExtInfo.[ReadOnly])
                    ' Total number of messages
                    Console.WriteLine("Total number of messages " + folderExtInfo.TotalMessageCount)
                Next

                'Disconnect to the remote IMAP server

                client.Disconnect()
            Catch ex As Exception
                System.Console.Write(Environment.NewLine + ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Getting folders information from IMAP server.")
        End Sub
    End Class
End Namespace
