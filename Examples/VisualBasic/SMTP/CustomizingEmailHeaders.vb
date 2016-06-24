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

    Public Class CustomizingEmailHeaders
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("MsgHeaders.msg")

            'Create an instance MailMessage class
            Dim msg As New MailMessage()

            'Specify ReplyTo
            msg.ReplyToList.Add("reply@reply.com")

            'From field
            msg.From = "sender@sender.com"

            'To field
            msg.[To].Add("receiver1@receiver.com")

            'Adding Cc and Bcc Addresses
            msg.CC.Add("receiver2@receiver.com")
            msg.Bcc.Add("receiver3@receiver.com")

            'Message subject
            msg.Subject = "test mail"

            'Specify Date
            msg.[Date] = New System.DateTime(2006, 3, 6)

            'Specify XMailer
            msg.XMailer = "Aspose.Email"

            'Specify Secret Header
            msg.Headers.Add("secret-header", "mystery")

            'Save message to disc
            msg.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode)

            Console.WriteLine(Environment.NewLine + "Message saved with customized headers successfully at " & dstEmail)
        End Sub
    End Class
End Namespace