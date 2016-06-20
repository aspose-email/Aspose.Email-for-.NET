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

    Public Class AddingNewMessage
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_IMAP()
            Dim dstEmail As String = dataDir & Convert.ToString("1234.eml")

            ' Create a message
            Dim msg As Aspose.Email.Mail.MailMessage
            msg = New Aspose.Email.Mail.MailMessage("user@domain1.com", "user@domain2.com", "subject", "message")

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
                ' Subscribe to the Inbox folder
                client.SelectFolder(ImapFolderInfo.InBox)
                client.SubscribeFolder(client.CurrentFolder.Name)

                ' Append the newly created message
                client.AppendMessage(client.CurrentFolder.Name, msg)

                System.Console.WriteLine("New Message Added Successfully")

                'Disconnect to the remote IMAP server

                client.Dispose()
            Catch ex As Exception
                System.Console.Write(Environment.NewLine + ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Added new message on IMAP server.")
        End Sub
    End Class
End Namespace
