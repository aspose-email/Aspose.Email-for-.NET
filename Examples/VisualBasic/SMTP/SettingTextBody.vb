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
    Public Class SettingTextBody
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("test.eml")

            'Declare msg as MailMessage instance
            Dim msg As New MailMessage()

            'use MailMessage properties like specify sender, recipient and message
            'use MailMessage properties like specify sender, recipient and message
            msg.From = "newcustomeronnet@gmail.com"
            msg.[To] = "newcustomeronnet2@gmail.com"
            msg.Subject = "Test subject"
            msg.Body = "This is text body"


            Dim client As SmtpClient = GetSmtpClient()

            Try
                'Client will send this message
                client.Send(msg)
                'Show only if message sent successfully
                Console.WriteLine("Message sent")

            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Email sent with Text body.")
        End Sub

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace