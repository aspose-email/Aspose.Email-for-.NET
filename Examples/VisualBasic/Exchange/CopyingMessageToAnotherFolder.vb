Imports System.Collections.Generic
Imports System.Configuration
Imports System.Net
Imports System.Threading
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class CopyingMessageToAnotherFolder
		Public Shared Sub Run()
			' ExStart:CopyingMessageToAnotherFolder
			Try
				' Create instance of EWSClient class by giving credentials
				Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
				Dim message As New MailMessage("from@domain.com", "to@domain.com", "EMAILNET-34997 - " + Guid.NewGuid().ToString(), "EMAILNET-34997 Exchange: Copy a message and get reference to the new Copy item")
				Dim messageUri As String = client.AppendMessage(message)
				Dim newMessageUri As String = client.CopyItem(messageUri, client.MailboxInfo.DeletedItemsUri)
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			End Try
			' ExEnd:CopyingMessageToAnotherFolder
		End Sub
	End Class
End Namespace