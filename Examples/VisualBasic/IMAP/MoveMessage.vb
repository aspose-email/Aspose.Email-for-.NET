Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class MoveMessage
        Private Shared Sub Run()
            ' ExStart: MoveMessage
            Using client As New ImapClient("host.domain.com", 993, "username", "password")
                Dim folderName As String = "EMAILNET-35151"
                If Not client.ExistFolder(folderName) Then
                    client.CreateFolder(folderName)
                End If
                Try
                    Dim message As New MailMessage("from@domain.com", "to@domain.com", "EMAILNET-35151 - " + Guid.NewGuid().ToString(), "EMAILNET-35151 ImapClient: Provide option to Move Message")
                    client.SelectFolder(ImapFolderInfo.InBox)
                    ' Append the new message to Inbox folder
                    Dim uniqueId As String = client.AppendMessage(ImapFolderInfo.InBox, message)
                    Dim messageInfoCol1 As ImapMessageInfoCollection = client.ListMessages()
                    Console.WriteLine(messageInfoCol1.Count)
                    ' Now move the message to the folder EMAILNET-35151
                    client.MoveMessage(uniqueId, folderName)
                    client.CommitDeletes()
                    ' Verify that the message was moved to the new folder
                    client.SelectFolder(folderName)
                    messageInfoCol1 = client.ListMessages()
                    Console.WriteLine(messageInfoCol1.Count)
                    ' Verify that the message was moved from the Inbox
                    client.SelectFolder(ImapFolderInfo.InBox)
                    messageInfoCol1 = client.ListMessages()
                    Console.WriteLine(messageInfoCol1.Count)
                Finally
                    Try
                        client.DeleteFolder(folderName)
                    Catch
                    End Try
                End Try
            End Using
            ' ExEnd: MoveMessage
        End Sub
    End Class
End Namespace