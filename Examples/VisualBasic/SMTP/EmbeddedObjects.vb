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

Namespace Aspose.Email.Examples.VisualBasic.Knowledge.SMTP

    Public Class EmbeddedObjects
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("EmbeddedImage.msg")

            'Create an instance of the MailMessage class
            Dim mail As New MailMessage()

            'Set the addresses
            mail.From = New MailAddress("test001@gmail.com")
            mail.[To].Add("test001@gmail.com")

            'Set the content
            mail.Subject = "This is an email"

            'Create the plain text part
            'It is viewable by those clients that don't support HTML
            Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString("This is my plain text content", Nothing, "text/plain")

            'Create the HTML part.
            'To embed images, we need to use the prefix 'cid' in the img src value.
            'The cid value will map to the Content-Id of a Linked resource.
            'Thus <img src='cid:barcode'> will map to a LinkedResource with a ContentId of //'barcode'.

            Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString("Here is an embedded image.<img src=cid:barcode>", Nothing, "text/html")

            'create the LinkedResource (embedded image)

            Dim barcode As New LinkedResource(dataDir & Convert.ToString("barcode.png"), MediaTypeNames.Image.Png)
            barcode.ContentId = "barcode"

            'Add the LinkedResource to the appropriate view

            mail.LinkedResources.Add(barcode)

            mail.AlternateViews.Add(plainView)
            mail.AlternateViews.Add(htmlView)

            mail.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode)

            Console.WriteLine(Environment.NewLine + "Message saved with embedded objects successfully at " & dstEmail)
        End Sub
    End Class
End Namespace