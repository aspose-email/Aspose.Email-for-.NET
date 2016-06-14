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
    Public Class SaveMessageAsDraft
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Dim dstEmail As String = dataDir & Convert.ToString("New-Draft.msg")

            ' Create a new instance of MailMessage class
            Dim message As New MailMessage()

            ' Set sender information
            message.From = "from@domain.com"

            ' Add recipients
            message.[To].Add("to1@domain.com")
            message.[To].Add("to2@domain.com")

            ' Set subject of the message
            message.Subject = "New message created by Aspose.Email"

            ' Set Html body of the message
            message.IsBodyHtml = True
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>"

            ' Create an instance of MapiMessage and load the MailMessag instance into it
            Dim mapiMsg As MapiMessage = MapiMessage.FromMailMessage(message)

            ' Set the MapiMessageFlags as UNSENT and FROMME
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT Or MapiMessageFlags.MSGFLAG_FROMME)

            ' Save the MapiMessage to disk
            mapiMsg.Save(dstEmail)

            Console.WriteLine(Environment.NewLine + "Created draft MSG at " & dstEmail)
        End Sub
    End Class
End Namespace