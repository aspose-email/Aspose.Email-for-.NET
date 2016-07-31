Imports System.Collections.Generic
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class SearchWithPagingSupport
        Private Shared Sub Run()
            ' ExStart:SearchWithPagingSupport
            Using client As New ImapClient("host.domain.com", 84, "username", "password")
                Try
                    ' Append some test messages
                    Dim messagesNum As Integer = 12
                    Dim itemsPerPage As Integer = 5
                    Dim message As MailMessage = Nothing
                    For i As Integer = 0 To messagesNum - 1
                        message = New MailMessage("from@domain.com", "to@domain.com", "EMAILNET-35128 - " + Guid.NewGuid().ToString(), "111111111111111")
                        client.AppendMessage(ImapFolderInfo.InBox, message)
                    Next
                    Dim body As String = "2222222222222"
                    For i As Integer = 0 To messagesNum - 1
                        message = New MailMessage("from@domain.com", "to@domain.com", "EMAILNET-35128 - " + Guid.NewGuid().ToString(), body)
                        client.AppendMessage(ImapFolderInfo.InBox, message)
                    Next

                    Dim iqb As New ImapQueryBuilder()
                    iqb.Body.Contains(body)
                    Dim query As MailQuery = iqb.GetQuery()

                    client.SelectFolder(ImapFolderInfo.InBox)
                    Dim totalMessageInfoCol As ImapMessageInfoCollection = client.ListMessages(query)
                    Console.WriteLine(totalMessageInfoCol.Count)

                    Dim pages As New List(Of ImapPageInfo)()
                    Dim pageInfo As ImapPageInfo = client.ListMessagesByPage(ImapFolderInfo.InBox, query, itemsPerPage)
                    pages.Add(pageInfo)
                    While Not pageInfo.LastPage
                        pageInfo = client.ListMessagesByPage(ImapFolderInfo.InBox, query, pageInfo.NextPage)
                        pages.Add(pageInfo)
                    End While
                    Dim retrievedItems As Integer = 0
                    For Each folderCol As ImapPageInfo In pages
                        retrievedItems += folderCol.Items.Count
                    Next
                Finally
                End Try
            End Using
            ' ExEnd: SearchWithPagingSupport
        End Sub
    End Class
End Namespace