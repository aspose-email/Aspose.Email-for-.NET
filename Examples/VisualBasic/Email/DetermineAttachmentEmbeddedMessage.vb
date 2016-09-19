Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Friend Class DetermineAttachmentEmbeddedMessage
        Public Shared Sub Run()
            ' ExStart:DetermineAttachmentEmbeddedMessage
            Dim dataDir As String = RunExamples.GetDataDir_Email() + "EmailWithAttandEmbedded.eml"
            Dim eml As MailMessage = MailMessage.Load(dataDir)

            If eml.Attachments(0).IsEmbeddedMessage Then
                Console.WriteLine("Attachment is an embedded message")
            End If
            ' ExEnd:DetermineAttachmentEmbeddedMessage
        End Sub
    End Class
End Namespace
