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

    Public Class LoadSmtpConfigFile
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("EmbeddedImage.msg")

            'Declare msg as MailMessage instance
            Dim msg As New MailMessage()

            'use MailMessage properties like specify sender, recipient and message
            msg.[To] = "asposetest123@gmail.com"
            msg.From = "aspose2@gmail.com"
            msg.Subject = "Test Email"
            msg.Body = "Hello World!"

            'Create an instance of SmtpClient class and load SMTP Authentication settings from Config file
            Dim client As New SmtpClient(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None))
            client.SecurityOptions = SecurityOptions.Auto

            Try
                'Client.Send will send this message
                client.Send(msg)
                'Message sent successfully
                System.Console.WriteLine("Message sent")

            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Message sent after loading configuration from config file.")
        End Sub
    End Class
End Namespace