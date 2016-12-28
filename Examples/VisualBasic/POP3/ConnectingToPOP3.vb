Imports Aspose.Email.Pop3

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.POP3
    Class ConnectingToPOP3
        Public Shared Sub Run()
            ' ExStart:GetServerExtensionsUsingPop3Client
            ' Connect and log in to POP3
            Dim client As New Pop3Client("pop.gmail.com", "username", "password")
            client.SecurityOptions = SecurityOptions.Auto
            client.Port = 993
            Dim getCaps As String() = client.GetCapabilities()
            For Each item As String In getCaps
                Console.WriteLine(item)
            Next
            ' ExEnd:GetServerExtensionsUsingPop3Client
        End Sub
    End Class
End Namespace
