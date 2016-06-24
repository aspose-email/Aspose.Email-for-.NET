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
Imports Aspose.Email.Outlook.Pst

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP

    Public Class MultipleRecipients
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("outputAttachments.msg")

            'Declare msg as MailMessage instance
            Dim message As New MailMessage()

            'use MailMessage properties like specify sender, recipient and message

            'Specify the recipients mail addresses
            message.[To].Add("receiver1@receiver.com")
            message.[To].Add("receiver2@receiver.com")
            message.[To].Add("receiver3@receiver.com")
            message.[To].Add("receiver4@receiver.com")

            message.CC.Add("CC1@receiver.com")
            message.CC.Add("CC2@receiver.com")

            message.Bcc.Add("Bcc1@receiver.com")
            message.Bcc.Add("Bcc2@receiver.com")

            message.From = "newcustomeronnet@gmail.com"
            message.Subject = "Test Email"
            message.Body = "Hello World!"


            'Create an instance of SmtpClient class
            Dim client As SmtpClient = GetSmtpClient()

            Try
                'Client will send this message
                client.Send(message)
                'Show only if message sent successfully
                Console.WriteLine("Message sent")

            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Email sent to multiple recipients successfully.")
        End Sub

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace