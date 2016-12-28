Imports Aspose.Email.Mail
Imports Aspose.Email.Mime

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP
    Class EmbeddedObjects
        Public Shared Sub Run()
            ' ExStart:EmbeddedObjects
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Dim dstEmail As String = dataDir & Convert.ToString("EmbeddedImage.msg")

            ' Create an instance of the MailMessage class and Set the addresses and Set the content
            Dim mail As New MailMessage()
            mail.From = New MailAddress("test001@gmail.com")
            mail.[To].Add("test001@gmail.com")
            mail.Subject = "This is an email"

            ' Create the plain text part It is viewable by those clients that don't support HTML
            Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString("This is my plain text content", Nothing, "text/plain")

            ' Create the HTML part.To embed images, we need to use the prefix 'cid' in the img src value.
            '            The cid value will map to the Content-Id of a Linked resource. Thus <img src='cid:barcode'> will map to a LinkedResource with a ContentId of //'barcode'. 

            Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString("Here is an embedded image.<img src=cid:barcode>", Nothing, "text/html")

            ' Create the LinkedResource (embedded image) and Add the LinkedResource to the appropriate view
            Dim barcode As New LinkedResource(dataDir & Convert.ToString("1.jpg"), MediaTypeNames.Image.Jpeg) With { _
                .ContentId = "barcode" _
            }
            mail.LinkedResources.Add(barcode)
            mail.AlternateViews.Add(plainView)
            mail.AlternateViews.Add(htmlView)
            mail.Save(dstEmail, SaveOptions.DefaultMsgUnicode)
            ' ExEnd:EmbeddedObjects
            Console.WriteLine(Convert.ToString(Environment.NewLine + "Message saved with embedded objects successfully at ") & dstEmail)
        End Sub
    End Class
End Namespace