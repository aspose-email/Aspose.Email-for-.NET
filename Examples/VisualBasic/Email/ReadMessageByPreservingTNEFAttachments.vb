Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class ReadMessageByPreservingTNEFAttachments
        Public Shared Sub Run()
            ' ExStart:ReadMessageByPreservingTNEFAttachments
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()

            Dim options As New MsgLoadOptions()
            options.PreserveTnefAttachments = True
            Dim eml As MailMessage = MailMessage.Load(dataDir & Convert.ToString("EmbeddedImage1.msg"), options)
            For Each attachment As Attachment In eml.Attachments
                Console.WriteLine(attachment.Name)
            Next
            ' ExEnd:ReadMessageByPreservingTNEFAttachments
        End Sub
    End Class
End Namespace