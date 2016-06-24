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

    Public Class MailMerge
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("EmbeddedImage.msg")

            'Create a new MailMessage instance
            Dim msg As New MailMessage()

            'Add subject and from address
            msg.Subject = "Hello, #FirstName#"
            msg.From = "sender@sender.com"

            'Add email address to send email
            msg.[To].Add("your.email@gmail.com")

            'Add mesage field to HTML body
            msg.HtmlBody = "Your message here"
            msg.HtmlBody += "Thank you for your interest in <STRONG>Aspose.Email</STRONG>."

            'Use GetSignment as the template routine, which will provide the same signature
            msg.HtmlBody += "<br><br>Have fun with it.<br><br>#GetSignature()#"

            'Create a new TemplateEngine with the MSG message.
            Dim engine As New TemplateEngine(msg)

            ' Register GetSignature routine. It will be used in MSG.
            engine.RegisterRoutine("GetSignature", New TemplateRoutine(AddressOf GetSignature))

            'Create an instance of DataTable
            'Fill a DataTable as data source
            Dim dt As New DataTable()
            dt.Columns.Add("Receipt", GetType(String))
            dt.Columns.Add("FirstName", GetType(String))
            dt.Columns.Add("LastName", GetType(String))

            'Create an instance of DataRow
            Dim dr As DataRow

            dr = dt.NewRow()
            dr("Receipt") = "abc<asposetest123@gmail.com>"
            dr("FirstName") = "a"
            dr("LastName") = "bc"
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Receipt") = "John<email.2@gmail.com>"
            dr("FirstName") = "John"
            dr("LastName") = "Doe"
            dt.Rows.Add(dr)

            dr = dt.NewRow()
            dr("Receipt") = "Third Recipient<email.3@gmail.com>"
            dr("FirstName") = "Third"
            dr("LastName") = "Recipient"
            dt.Rows.Add(dr)

            Dim messages As MailMessageCollection
            Try
                'Create messages from the message and datasource.
                messages = engine.Instantiate(dt)

                'Create an instance of SmtpClient and specify server, port, username and password
                Dim client As SmtpClient = GetSmtpClient()

                'Send messages in bulk
                client.Send(messages)
            Catch ex As MailException
                System.Diagnostics.Debug.WriteLine(ex.ToString())

            Catch ex As SmtpException
                System.Diagnostics.Debug.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Message sent after performing mail merge.")
        End Sub

        'Template routine to provide signature
        Private Shared Function GetSignature(args As Object()) As Object
            Return "Aspose.Email Team<br>Aspose Ltd.<br>" + DateTime.Now.ToShortDateString()
        End Function

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace