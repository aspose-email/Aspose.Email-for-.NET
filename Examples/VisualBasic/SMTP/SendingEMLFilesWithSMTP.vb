Imports System.IO
Imports Aspose.Email.Mail

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP

    Public Class SendingEMLFilesWithSMTP
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("test-load.eml")

            'Create an instance of the MailMessage class
            Dim msg As New MailMessage()

            'Import from EML format
            msg = MailMessage.Load(dstEmail, New EmlLoadOptions())

            'Create an instance of SmtpClient class
            Dim client As SmtpClient = GetSmtpClient()

            Try
                'Client.Send will send this message
                client.Send(msg)
                System.Console.WriteLine("Message sent")

            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Convert.ToString(Environment.NewLine + "Email sent using EML file successfully. ") & dstEmail)
        End Sub

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace