 Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class LoadMSGFiles

        Public Shared Sub Run()
            ' ExStart:LoadMSGFiles
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create an instance of MapiMessage from file
            Dim msg As MapiMessage = MapiMessage.FromFile(dataDir & Convert.ToString("message.msg"))

            ' Get subject
            Console.WriteLine("Subject:" + msg.Subject)

            ' Get from address
            Console.WriteLine("From:" + msg.SenderEmailAddress)

            ' Get body
            Console.WriteLine("Body" + msg.Body)

            ' Get recipients information
            Console.WriteLine("Recipient: " + msg.Recipients.ToString())

            ' Get attachments
            For Each att As MapiAttachment In msg.Attachments
                Console.Write("Attachment Name: " + att.FileName)
                Console.Write("Attachment Display Name: " + att.DisplayName)
            Next
            ' ExEnd:LoadMSGFiles
        End Sub
    End Class
End Namespace