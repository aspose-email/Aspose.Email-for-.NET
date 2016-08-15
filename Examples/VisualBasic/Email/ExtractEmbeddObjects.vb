Imports System.IO
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class ExtractEmbeddObjects
        Public Shared Sub Run()
            ' ExStart:CreatingTNEFFromMSG
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()

            ' Create an instance of MailMessage and load an email file
            Dim mailMsg As MailMessage = MailMessage.Load(dataDir & Convert.ToString("EmailWithAttandEmbedded.eml"))

            ' Extract Attachments from the message
            For Each attachment As Attachment In mailMsg.Attachments
                ' To display the the attachment file name
                Console.WriteLine(attachment.Name)

                ' Save the attachment to disc
                attachment.Save(attachment.Name)

                ' You can also save the attachment to memory stream
                Dim memorystream As New MemoryStream()

                attachment.Save(dataDir & memorystream.ToString() & "_out")
            Next

            ' Extract Inline images from the message
            For Each lr As LinkedResource In mailMsg.LinkedResources
                Console.WriteLine(lr.ContentType.Name)

                lr.Save(lr.ContentType.Name)
            Next
        End Sub
    End Class
End Namespace
