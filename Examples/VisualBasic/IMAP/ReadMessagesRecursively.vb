Imports System.IO
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class ReadMessagesRecursively
        ' ExStart:ReadMessagesRecursively
        Public Shared Sub Run()
            ' Create an instance of the ImapClient class
            Dim client As New ImapClient()

            ' Specify host, username, password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com"
            client.Username = "your.username@gmail.com"
            client.Password = "your.password"
            client.Port = 993
            client.SecurityOptions = SecurityOptions.Auto
            Try
                ' The root folder (which will be created on disk) consists of host and username
                Dim rootFolder As String = client.Host + "-" + client.Username

                ' Create the root folder and List all the folders from IMAP server
                Directory.CreateDirectory(rootFolder)
                Dim folderInfoCollection As ImapFolderInfoCollection = client.ListFolders()
                For Each folderInfo As ImapFolderInfo In folderInfoCollection
                    ' Call the recursive method to read messages and get sub-folders
                    ListMessagesInFolder(folderInfo, rootFolder, client)
                Next
                ' Disconnect to the remote IMAP server
                client.Dispose()
            Catch ex As Exception
                Console.Write(Environment.NewLine + ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Downloaded messages recursively from IMAP server.")
        End Sub


        ' Recursive method to get messages from folders and sub-folders
        Private Shared Sub ListMessagesInFolder(folderInfo As ImapFolderInfo, rootFolder As String, client As ImapClient)
            ' Create the folder in disk (same name as on IMAP server)
            Dim currentFolder As String = RunExamples.GetDataDir_IMAP()
            Directory.CreateDirectory(currentFolder)

            ' Read the messages from the current folder, if it is selectable
            If folderInfo.Selectable Then
                ' Send status command to get folder info
                Dim folderInfoStatus As ImapFolderInfo = client.GetFolderInfo(folderInfo.Name)
                Console.WriteLine(folderInfoStatus.Name + " folder selected. New messages: " + folderInfoStatus.NewMessageCount + ", Total messages: " + folderInfoStatus.TotalMessageCount)

                ' Select the current folder and List messages
                client.SelectFolder(folderInfo.Name)
                Dim msgInfoColl As ImapMessageInfoCollection = client.ListMessages()
                Console.WriteLine("Listing messages....")
                For Each msgInfo As ImapMessageInfo In msgInfoColl
                    ' Get subject and other properties of the message
                    Console.WriteLine("Subject: " + msgInfo.Subject)
                    Console.WriteLine("Read: " + msgInfo.IsRead + ", Recent: " + msgInfo.Recent + ", Answered: " + msgInfo.Answered)

                    ' Get rid of characters like ? and :, which should not be included in a file name and Save the message in MSG format
                    Dim fileName As String = msgInfo.Subject.Replace(":", " ").Replace("?", " ")
                    Dim msg As MailMessage = client.FetchMessage(msgInfo.SequenceNumber)
                    msg.Save((Convert.ToString(currentFolder & Convert.ToString("\")) & fileName) + "-" + msgInfo.SequenceNumber + ".msg", SaveOptions.DefaultMsgUnicode)
                Next
                Console.WriteLine("============================" & vbLf)
            Else
                Console.WriteLine(folderInfo.Name + " is not selectable.")
            End If

            Try
                ' If this folder has sub-folders, call this method recursively to get messages
                Dim folderInfoCollection As ImapFolderInfoCollection = client.ListFolders(folderInfo.Name)
                For Each subfolderInfo As ImapFolderInfo In folderInfoCollection
                    ListMessagesInFolder(subfolderInfo, rootFolder, client)
                Next
            Catch generatedExceptionName As Exception
            End Try
            ' ExEnd:ReadMessagesRecursively
        End Sub
    End Class
End Namespace