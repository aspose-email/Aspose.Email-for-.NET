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

    Public Class CustomizingEmailHeader
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("test.eml")

            'Create an instance MailMessage class
            Dim msg As New MailMessage()

            'Specify ReplyTo
            msg.ReplyToList.Add("reply@reply.com")

            'From field
            msg.From = "sender@sender.com"

            'To field
            msg.[To].Add("receiver1@receiver.com")

            'Adding CC and BCC Addresses
            msg.CC.Add("receiver2@receiver.com")
            msg.Bcc.Add("receiver3@receiver.com")

            'Message subject
            msg.Subject = "test mail"

            'Specify Date
            msg.[Date] = New System.DateTime(2006, 3, 6)

            'Specify XMailer
            msg.XMailer = "Aspose.Email"

            msg.Headers.Add_("secret-header", "mystery")

            'Create an instance of SmtpClient class
            Dim client As SmtpClient = GetSmtpClient()

            Try
                'Client.Send will send this message
                client.Send(msg)
                'Message sent successfully
                System.Console.WriteLine("Message sent")

            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Email sent with customized headers.")
        End Sub

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace