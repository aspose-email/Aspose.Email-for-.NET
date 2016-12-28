Imports Aspose.Email.Imap
Imports Aspose.Email.Protocols.Proxy

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class ProxyServerAccessMailbox
		Public Shared Sub Run()
			' ExStart:ProxyServerAccessMailbox
			' Connect and log in to IMAP and set SecurityOptions
			Dim client As New ImapClient("imap.domain.com", "username", "password")
			client.SecurityOptions = SecurityOptions.Auto
			Dim proxyAddress As String = "192.168.203.142"
			' proxy address
			Dim proxyPort As Integer = 1080
			' proxy port
			Dim proxy As New SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5)

			' Set the proxy
			client.SocksProxy = proxy
			client.SelectFolder("Inbox")
			' ExEnd:ProxyServerAccessMailbox
		End Sub
	End Class
End Namespace