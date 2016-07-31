Imports Aspose.Email.Imap

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class RetreivingServerExtensions
		Public Shared Sub Run()
			' ExStart:RetreivingServerExtensions
			' Connect and log in to IMAP
			Dim client As New ImapClient("imap.gmail.com", "username", "password")
			Dim getCapabilities As String() = client.GetCapabilities()
			For Each getCap As String In getCapabilities
				Console.WriteLine(getCap)
			Next
			' ExEnd:RetreivingServerExtensions
		End Sub
	End Class
End Namespace