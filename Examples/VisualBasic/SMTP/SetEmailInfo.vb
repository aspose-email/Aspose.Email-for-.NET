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

    Public Class SetEmailInfo
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("test.eml")

            'Create an instance MailMessage class
            Dim msg As New MailMessage()

            'Setting Date 
            msg.[Date] = DateTime.Now

            'Setting Priority
            msg.Priority = MailPriority.High

            'Setting Sensitivity
            msg.Sensitivity = MailSensitivity.Normal

            'use MailMessage properties like specify sender, recipient and message
            msg.[To] = "asposetest123@gmail.com"
            msg.From = "asposetest123@gmail.com"
            msg.Subject = "Test Email"
            msg.Body = "Hello World!"


            'Create an instance of SmtpClient class
            Dim client As SmtpClient = GetSmtpClient()

            Try
                'Client.Send will send this message
                client.Send(msg)
                'Message sent successfully
                Console.WriteLine("Message sent")

            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Email sent with setting message properties.")
        End Sub

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace