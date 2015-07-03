' ///////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
' ///////////////////////////////////////////////////////////////////////

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

Public Class SaveExchangeTaskToDisc
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_Exchange()
        Dim dstEmail As String = dataDir & Convert.ToString("Message.msg")

        Dim task As New ExchangeTask()
        task.Subject = "TASK-ID - " + Guid.NewGuid().ToString()
        task.Status = ExchangeTaskStatus.InProgress
        task.StartDate = DateTime.Now
        task.DueDate = task.StartDate.AddDays(3)
        task.Save(dstEmail)

        Console.WriteLine(Convert.ToString(Environment.NewLine + "Task saved on disk successfully at ") & dstEmail)
    End Sub
End Class
