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
    Public Class SSLEnabledSMTPServer
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("test-load.eml")

            Dim client As New Aspose.Email.Mail.SmtpClient("smtp.gmail.com")

            ' Set username
            client.Username = "your.email@gmail.com"

            ' Set password
            client.Password = "your.password"

            ' Set the port to 587. This is the SSL port of SMTP server
            client.Port = 587

            ' Set the security mode to explicit
            client.SecurityOptions = SecurityOptions.Auto

            'Declare msg as MailMessage instance
            Dim msg As New MailMessage()

            'use MailMessage properties like specify sender, recipient and message
            msg.[To] = "newcustomeronnet@gmail.com"
            msg.From = "newcustomeronnet@gmail.com"
            msg.Subject = "Test Email"
            msg.Body = "Hello World!"

            Try
                'Client will send this message
                client.Send(msg)
                'Show only if message sent successfully
                System.Console.WriteLine("Message sent")

            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Email sent SSL successfully.")
        End Sub
    End Class
End Namespace