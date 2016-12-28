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
    Public Class CreateExchangeTask
        Public Shared Sub Run()
            ' Create instance of EWSClient class by giving credentials
            Dim client As IEWSClient = EWSClient.GetEWSClient("https:outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

            ' Create Exchange task object
            Dim task As New ExchangeTask()

            ' Set task subject
            task.Subject = "New-Test"

            ' Set task status to In progress
            task.Status = ExchangeTaskStatus.InProgress

            ' Create task on exchange
            client.CreateTask(client.MailboxInfo.TasksUri, task)

            Console.WriteLine(Environment.NewLine + "Task created on Exchange Server successfully.")
        End Sub
    End Class
End Namespace