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

Namespace Aspose.Email.Examples.VisualBasic.Email
    Public Class ExtractingEmailHeaders
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Dim dstEmail As String = dataDir & Convert.ToString("test.eml")

            Dim message As MailMessage

            'Create MailMessage instance by loading an EML file
            message = MailMessage.Load(dataDir & Convert.ToString("email-headers.eml"), New EmlLoadOptions())

            Console.WriteLine(vbLf & vbLf & "headers:" & vbLf & vbLf)

            'print out all the headers
            Dim index As Integer = 0
            For Each header As [String] In message.Headers
                Console.Write(header + " - ")
                '.GetValues(header).Length.ToString())
                Console.WriteLine(message.Headers.Get(index))
                index = index + 1
            Next

            Console.WriteLine(Environment.NewLine + "Displayed email headers from " & dstEmail)
        End Sub
    End Class
End Namespace