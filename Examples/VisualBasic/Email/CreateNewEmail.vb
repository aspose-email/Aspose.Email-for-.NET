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

    Public Class CreateNewEmail
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Dim dstEmail As String = dataDir & Convert.ToString("test.eml")

            ' Create a new instance of MailMessage class
            Dim message As New MailMessage()

            ' Set subject of the message
            message.Subject = "New message created by Aspose.Email for .NET"

            ' Set Html body
            message.IsBodyHtml = True
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>"

            ' Set sender information
            message.From = "from@domain.com"

            ' Add TO recipients
            message.[To].Add("to1@domain.com")
            message.[To].Add("to2@domain.com")

            'Add CC recipients
            message.CC.Add("cc1@domain.com")
            message.CC.Add("cc2@domain.com")

            ' Save message in EML, MSG and MHTML formats
            message.Save(dataDir & Convert.ToString("Message.eml"), Aspose.Email.Mail.SaveOptions.DefaultEml)
            message.Save(dataDir & Convert.ToString("Message.msg"), Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode)
            message.Save(dataDir & Convert.ToString("Message.mhtml"), Aspose.Email.Mail.SaveOptions.DefaultMhtml)

            Console.WriteLine(Environment.NewLine + "Created new email in EML, MSG and MHTML formats successfully.")
        End Sub
    End Class
End Namespace