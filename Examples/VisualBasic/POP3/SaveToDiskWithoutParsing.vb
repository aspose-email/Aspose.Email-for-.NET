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

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
    Public Class SaveToDiskWithoutParsing
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_POP3()
            Dim dstEmail As String = dataDir & Convert.ToString("1234.eml")

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
                'Save message to disk by message sequence number
                client.SaveMessage(1, dstEmail)

                client.Dispose()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            Console.WriteLine(Environment.NewLine + "Retrieved email messages using POP3 ")
        End Sub
    End Class
End Namespace