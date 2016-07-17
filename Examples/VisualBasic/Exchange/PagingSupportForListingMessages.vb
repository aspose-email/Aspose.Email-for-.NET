Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange

    Public Class PagingSupportForListingMessages
        Private Shared Sub Run()
            'ExStart: PagingSupportForListingMessages
            '<summary>
            'This example shows how to list messages from Exchange Server with Paging support
            'Introduced in Aspose.Email for .NET 6.4.0
            'ExchangeMessagePageInfo ListMessagesByPage(string folder, int itemsPerPage)
            'ExchangeMessagePageInfo ListMessagesByPage(string folder, int itemsPerPage, int offset)
            'ExchangeMessagePageInfo ListMessagesByPage(string folder, int itemsPerPage, int pageOffset, ExchangeListMessagesOptions options)
            'ExchangeMessagePageInfo ListMessagesByPage(string folder, PageInfo pageInfo)
            'ExchangeMessagePageInfo ListMessagesByPage(string folder, PageInfo pageInfo, ExchangeListMessagesOptions options)
            '</summary>
            Using client As IEWSClient = EWSClient.GetEWSClient("exchange.domain.com", "username", "password")
                Try
                    'Create some test messages to be added to server for retrieval later
                    Dim messagesNum As Integer = 12
                    Dim itemsPerPage As Integer = 5
                    Dim message As MailMessage = Nothing
                    For i As Integer = 0 To messagesNum - 1
                        message = New MailMessage("from@domain.com", "to@domain.com", "EMAILNET-35157_1 - " + Guid.NewGuid().ToString(), "EMAILNET-35157 Move paging parameters to separate class")
                        client.AppendMessage(client.MailboxInfo.InboxUri, message)
                    Next
                    'Verfiy that the messages have been added to the server
                    Dim totalMessageInfoCol As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri)
                    Console.WriteLine(totalMessageInfoCol.Count)

                    '/ RETREIVING THE MESSAGES USING PAGING SUPPORT 

                    Dim pages As New List(Of ExchangeMessagePageInfo)()
                    Dim pageInfo As ExchangeMessagePageInfo = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage)
                    'Total Pages Count
                    Console.WriteLine(pageInfo.TotalCount)

                    pages.Add(pageInfo)
                    While Not pageInfo.LastPage
                        pageInfo = client.ListMessagesByPage(client.MailboxInfo.InboxUri, itemsPerPage, pageInfo.PageOffset + 1)
                        pages.Add(pageInfo)
                    End While
                    Dim retrievedItems As Integer = 0
                    For Each pageCol As ExchangeMessagePageInfo In pages
                        retrievedItems += pageCol.Items.Count
                    Next
                    'Verify total message count using paging
                    Console.WriteLine(retrievedItems)
                Finally
                End Try
            End Using
            'ExEnd: PagingSupportForListingMessages

        End Sub
    End Class
End Namespace