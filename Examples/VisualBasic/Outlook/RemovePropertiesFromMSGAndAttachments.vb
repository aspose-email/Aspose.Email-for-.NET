Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class RemovePropertiesFromMSGAndAttachments
        Public Shared Sub Run()
            ' ExStart:RemovePropertiesFromMSGAndAttachments
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim mapi As New MapiMessage("from@doamin.com", "to@domain.com", "subject", "body")
            mapi.SetBodyContent("<html><body><h1>This is the body content</h1></body></html>", BodyContentType.Html)
            Dim attachment As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("message.msg"))
            mapi.Attachments.Add("Outlook2 Test subject.msg", attachment)
            Console.WriteLine("Before removal " + mapi.Attachments(mapi.Attachments.Count - 1).Properties.Count.ToString())

            mapi.Attachments(mapi.Attachments.Count - 1).RemoveProperty(923467779)
            ' Delete anyone property
            Console.WriteLine("After removal = " + mapi.Attachments(mapi.Attachments.Count - 1).Properties.Count.ToString())
            mapi.Save("EMAIL_589265.msg")

            Dim mapi2 As MapiMessage = MapiMessage.FromFile("EMAIL_589265.msg")
            Console.WriteLine("Reloaded = " + mapi2.Attachments(mapi2.Attachments.Count - 1).Properties.Count.ToString())
            ' ExEnd:RemovePropertiesFromMSGAndAttachments
        End Sub
    End Class
End Namespace
