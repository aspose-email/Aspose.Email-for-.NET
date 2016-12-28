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
Imports System.Net

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Public Class UpdateExchangeTask
        Public Shared Sub Run()
            ' Create and initialize credentials
            Dim credentials = New NetworkCredential("username", "12345")

            ' Create instance of ExchangeClient class by giving credentials
            Dim client As IEWSClient = EWSClient.GetEWSClient("https:outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

            ' Get all tasks info collection from exchange
            Dim tasks As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.TasksUri)

            ' Parse all the tasks info in the list
            For Each info As ExchangeMessageInfo In tasks
                ' Fetch task from exchange using current task info
                Dim task As ExchangeTask = client.FetchTask(info.UniqueUri)

                ' Update the task status to NotStarted
                task.Status = ExchangeTaskStatus.NotStarted

                ' Set the task due date
                task.DueDate = New DateTime(2013, 2, 26)

                ' Set task priority
                task.Priority = MailPriority.Low

                ' Update task on exchange
                client.UpdateTask(task)
            Next

            Console.WriteLine(Environment.NewLine + "Task updated on Exchange Server successfully.")
        End Sub
    End Class
End Namespace
