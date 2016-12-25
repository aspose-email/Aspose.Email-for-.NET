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

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Public Class FilterTasksFromServer
        Public Shared Sub Run()
            ' ExStart:FilterTasksFromServer
            ' Set mailboxURI, Username, password, domain information
            Dim mailboxUri As String = "https://ex2010/ews/exchange.asmx"
            Dim username As String = "test.exchange"
            Dim password As String = "pwd"
            Dim domain As String = "ex2010.local"
            Dim credentials As New NetworkCredential(username, password, domain)
            Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credentials)

            Dim queryBuilder As ExchangeQueryBuilder = Nothing
            Dim query As MailQuery = Nothing
            Dim fetchedTask As ExchangeTask = Nothing
            Dim messageInfoCol As ExchangeMessageInfoCollection = Nothing
            client.TimezoneId = "Central Europe Standard Time"
            Dim values As Array = [Enum].GetValues(GetType(ExchangeTaskStatus))

            'Now retrieve the tasks with specific statuses
            For Each status As ExchangeTaskStatus In values
                queryBuilder = New ExchangeQueryBuilder()
                queryBuilder.TaskStatus.Equals(status)
                query = queryBuilder.GetQuery()
                messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query)
                fetchedTask = client.FetchTask(messageInfoCol(0).UniqueUri)
            Next

            'retrieve all other than specified
            For Each status As ExchangeTaskStatus In values
                queryBuilder = New ExchangeQueryBuilder()
                queryBuilder.TaskStatus.NotEquals(status)
                query = queryBuilder.GetQuery()
                messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query)
            Next

            'specifying multiple criterion
            Dim selectedStatuses As ExchangeTaskStatus() = New ExchangeTaskStatus() {ExchangeTaskStatus.Completed, ExchangeTaskStatus.InProgress}
            queryBuilder = New ExchangeQueryBuilder()
            queryBuilder.TaskStatus.[In](selectedStatuses)
            query = queryBuilder.GetQuery()
            messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query)

            queryBuilder = New ExchangeQueryBuilder()
            queryBuilder.TaskStatus.NotIn(selectedStatuses)
            query = queryBuilder.GetQuery()
            messageInfoCol = client.ListMessages(client.MailboxInfo.TasksUri, query)
            'ExEnd:FilterTasksFromServer
        End Sub
    End Class
End Namespace
