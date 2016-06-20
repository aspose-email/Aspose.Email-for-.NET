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

Namespace Aspose.Email.Examples.VisualBasic.Knowledge.POP3
    Public Class RetrievingEmailMessages
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_POP3()
            Dim dstEmail As String = dataDir & Convert.ToString("message.msg")

            'Create an instance of the Pop3Client class
            Dim client As New Pop3Client()

            'Specify host, username and password for your client
            client.Host = "pop.gmail.com"

            ' Set username
            client.Username = "your.username@gmail.com"

            ' Set password
            client.Password = "your.password"

            ' Set the port to 995. This is the SSL port of POP3 server
            client.Port = 995

            ' Enable SSL
            client.SecurityOptions = SecurityOptions.Auto

            Try
                Dim messageCount As Integer = client.GetMessageCount()
                ' Create an instance of the MailMessage class
                Dim msg As MailMessage

                For i As Integer = 1 To messageCount
                    'Retrieve the message in a MailMessage class
                    msg = client.FetchMessage(i)

                    Console.WriteLine("From:" + msg.From.ToString())
                    Console.WriteLine("Subject:" + msg.Subject)
                    Console.WriteLine(msg.HtmlBody)

                Next
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            Finally
                client.Dispose()
            End Try

            Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ")
        End Sub
    End Class
End Namespace