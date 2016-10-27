Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ParseOutlookMessageFile
        Public Shared Sub Run()
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load Microsoft Outlook email message file
            Dim msg As MapiMessage = MapiMessage.FromMailMessage(dataDir & Convert.ToString("Message.eml"))

            ' Obtain subject of the email message, SenderName, body and attachments count
            Console.WriteLine("Subject:" + msg.Subject.ToString())
            Console.WriteLine("From:" + msg.SenderName.ToString())
            Console.WriteLine("Body:" + msg.Body.ToString())
            Console.WriteLine("Attachment Count:" + msg.Attachments.Count.ToString())

            ' Iterate through the attachments
            For Each attachment As MapiAttachment In msg.Attachments
                Console.WriteLine("Attachment:" + attachment.FileName)
                attachment.Save(attachment.LongFileName.ToString())
            Next
        End Sub
    End Class
End Namespace