Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ReadingNamedMAPIPropertyFromAttachment
        Public Shared Sub Run()
            Dim getProperty As String = GetNamedPropertyByAspose()
            Console.WriteLine(Convert.ToString("getProperty") & getProperty)
        End Sub

        ' Gets attachment's named property
        ' Property name is "CustomAttGuid"
        Private Shared Function GetNamedPropertyByAspose() As String
            'ExStart:ReadingNamedMAPIPropertyFromAttachment
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load from file
            Dim mail As MailMessage = MailMessage.Load(dataDir & Convert.ToString("outputAttachments.msg"))

            Dim mapi = MapiMessage.FromMailMessage(mail)

            For Each namedProperty As MapiNamedProperty In mapi.Attachments(0).NamedProperties.Values
                If String.Compare(namedProperty.NameId, "CustomAttGuid", StringComparison.OrdinalIgnoreCase) = 0 Then

                    Return namedProperty.GetString()
                End If
            Next
            Return String.Empty
            'ExEnd:ReadingNamedMAPIPropertyFromAttachment
        End Function
    End Class
End Namespace
