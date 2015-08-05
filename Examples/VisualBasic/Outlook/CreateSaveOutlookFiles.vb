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

Public Class CreateSaveOutlookFiles
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_Outlook()
        Dim dst As String = dataDir & Convert.ToString("message.msg")

        ' Create an instance of MailMessage class
        Dim mailMsg As New MailMessage()

        ' Set FROM field of the message
        mailMsg.From = "from@domain.com"

        ' Set TO field of the message
        mailMsg.[To].Add("to@domain.com")

        ' Set SUBJECT of the message
        mailMsg.Subject = "creating an outlook message file"

        ' Set BODY of the message
        mailMsg.Body = "This message is created by Aspose.Email"

        ' Create an instance of MapiMessage class and pass MailMessage as argument
        Dim outlookMsg As MapiMessage = MapiMessage.FromMailMessage(mailMsg)

        ' Save the message (msg) file
        outlookMsg.Save(dst)

        Console.WriteLine(Environment.NewLine + "MSG saved successfully at " & dst)
    End Sub
End Class
