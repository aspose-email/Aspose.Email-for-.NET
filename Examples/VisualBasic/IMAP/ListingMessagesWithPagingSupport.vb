Imports System.Collections.Generic
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class ListingMessagesWithPagingSupport
        Private Shared Sub Run()
            ' ExStart: ListingMessagesWithPagingSupport
            Using client As New ImapClient("host.domain.com", 993, "username", "password")
                Try
                    Dim messagesNum As Integer = 12
                    Dim itemsPerPage As Integer = 5
                    Dim message As MailMessage = Nothing
                    ' Create some test messages and append these to server's inbox
                    For i As Integer = 0 To messagesNum - 1
                        message = New MailMessage("from@domain.com", "to@domain.com", "EMAILNET-35157 - " + Guid.NewGuid().ToString(), "EMAILNET-35157 Move paging parameters to separate class")
                        client.AppendMessage(ImapFolderInfo.InBox, message)
                    Next

                    ' List messages from inbox
                    client.SelectFolder(ImapFolderInfo.InBox)
                    Dim totalMessageInfoCol As ImapMessageInfoCollection = client.ListMessages()
                    ' Verify the number of messages added
                    Console.WriteLine(totalMessageInfoCol.Count)

                    ' RETREIVE THE MESSAGES USING PAGING SUPPORT

                    Dim pages As New List(Of ImapPageInfo)()
                    Dim pageInfo As ImapPageInfo = client.ListMessagesByPage(itemsPerPage)
                    Console.WriteLine(pageInfo.TotalCount)
                    pages.Add(pageInfo)
                    While Not pageInfo.LastPage
                        pageInfo = client.ListMessagesByPage(pageInfo.NextPage)
                        pages.Add(pageInfo)
                    End While
                    Dim retrievedItems As Integer = 0
                    For Each folderCol As ImapPageInfo In pages
                        retrievedItems += folderCol.Items.Count
                    Next
                    Console.WriteLine(retrievedItems)
                Finally
                End Try
            End Using
            ' ExEnd: ListingMessagesWithPagingSupport
        End Sub
    End Class
End Namespace