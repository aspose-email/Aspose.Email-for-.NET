'''///////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'''///////////////////////////////////////////////////////////////////////
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports System.IO
Imports System.Net

Module Module1

    Sub Main()
        'Initialize your license here

        'now call any of the methods below as per your requirements. Please note that you need a valid Exchange URI, 
        'UserName, and password (and domain if applicable) to try these examples.

    End Sub

    Private Sub CreateExchangeTask()
        ' Create instance of EWSClient class by giving credentials
        Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

        ' Create Exchange task object
        Dim task As New ExchangeTask()

        ' Set task subject
        task.Subject = "New-Test"

        ' Set task status to In progress
        task.Status = ExchangeTaskStatus.InProgress

        ' Create task on exchange
        client.CreateTask(client.MailboxInfo.TasksUri, task)
    End Sub

    Private Sub UpdateExchangeTask()
        ' Create and initialize credentials
        Dim credentials = New NetworkCredential("username", "12345")

        ' Create instance of ExchangeClient class by giving credentials
        Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

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
    End Sub

    Private Sub DeleteExchangeTask()
        ' Create instance of ExchangeClient class by giving credentials
        Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

        ' Get all tasks info collection from exchange
        Dim tasks As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.TasksUri)

        ' Parse all the tasks info in the list
        For Each info As ExchangeMessageInfo In tasks
            ' Fetch task from exchange using current task info
            Dim task As ExchangeTask = client.FetchTask(info.UniqueUri)

            ' Check if the current task fulfills the search criteria
            If task.Subject.Equals("test") Then
                'Delete task from exchange
                client.DeleteTask(task.UniqueUri, DeleteTaskOptions.DeletePermanently)
            End If
        Next
    End Sub

    Private Sub SendExchangeTask()
        ' Create instance of ExchangeClient class by giving credentials
        Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

        Dim loadOptions As New MailMessageLoadOptions()
        loadOptions.MessageFormat = MessageFormat.Msg
        loadOptions.FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments

        ' load task from .msg file
        Dim eml As MailMessage = MailMessage.Load("task.msg", loadOptions)
        eml.From = "firstname.lastname@domain.com"
        eml.[To].Clear()
        eml.[To].Add(New Aspose.Email.Mail.MailAddress("firstname.lastname@domain.com"))

        client.Send(eml)
    End Sub

    Private Sub SpecifyTimeZoneForTask()
        Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
        client.TimezoneId = "Central Europe Standard Time"
    End Sub

    'Available for Aspose.Email for .NET 5.4.0 and onwards
    Private Sub SaveExchangeTaskToDisc()
        Dim dataDir As String = Path.GetFullPath("../../../Data/")

        Dim task As New ExchangeTask()
        task.Subject = "EMAILNET-34759 - " + Guid.NewGuid().ToString()
        task.Status = ExchangeTaskStatus.InProgress
        task.StartDate = New DateTime(2028, 10, 6, 12, 30, 0)
        task.DueDate = task.StartDate.AddDays(3)
        task.Save(dataDir & Convert.ToString("ExchangeTask.msg"))
    End Sub
End Module
