Imports Aspose.Email.Mail
Imports Aspose.Email.Protocols.Proxy

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP
    Class SendEmailViaProxyServer
        Public Shared Sub Run()
            ' ExStart:SendEmailViaProxyServer
            Dim client As New SmtpClient("smtp.domain.com", "username", "password")
            client.SecurityOptions = SecurityOptions.SSLImplicit
            Dim proxyAddress As String = "192.168.203.142"
            ' proxy address
            Dim proxyPort As Integer = 1080
            ' proxy port
            Dim proxy As New SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5)
            client.SocksProxy = proxy
            client.Send(New MailMessage("sender@domain.com", "receiver@domain.com", "Sending Email via proxy", "Implement socks proxy protocol for versions 4, 4a, 5 (only Username/Password authentication)"))
            ' ExEnd:SendEmailViaProxyServer
        End Sub
    End Class
End Namespace