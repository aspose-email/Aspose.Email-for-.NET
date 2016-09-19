Imports Aspose.Email.Pop3

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
    Class RetrieveMessageSummaryInformationUsingUniqueId
        Public Shared Sub Run()
            ' ExStart:RetrieveMessageSummaryInformationUsingUniqueId
            Dim uniqueId As String = "unique id of a message from server"
            Dim client As New Pop3Client("host.domain.com", 456, "username", "password")
            client.SecurityOptions = SecurityOptions.Auto
            Dim messageInfo As Pop3MessageInfo = client.GetMessageInfo(uniqueId)

            If messageInfo IsNot Nothing Then
                Console.WriteLine(messageInfo.Subject)
                Console.WriteLine(messageInfo.[Date])
            End If
            ' ExEnd:RetrieveMessageSummaryInformationUsingUniqueId
        End Sub
    End Class
End Namespace
