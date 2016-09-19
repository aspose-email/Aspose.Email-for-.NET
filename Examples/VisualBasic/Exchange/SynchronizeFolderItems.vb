Imports System.Collections.Generic
Imports System.Net
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
	Class SynchronizeFolderItems
		Public Shared Sub Run()
			' ExStart:SynchronizeFolderItems
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

			Dim message1 As New MailMessage("user@domain.com", "user@domain.com", "EMAILNET-34738 - " + Guid.NewGuid().ToString(), "EMAILNET-34738 Sync Folder Items")
			client.Send(message1)

			Dim message2 As New MailMessage("user@domain.com", "user@domain.com", "EMAILNET-34738 - " + Guid.NewGuid().ToString(), "EMAILNET-34738 Sync Folder Items")
			client.Send(message2)

			Dim messageInfoCol As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri)

            Dim result As SyncFolderResult = client.SyncFolder(client.MailboxInfo.InboxUri, CType(Nothing, SyncFolderType))
			Console.WriteLine(result.NewItems.Count)
			Console.WriteLine(result.ChangedItems.Count)
			Console.WriteLine(result.ReadFlagChanged.Count)
			Console.WriteLine(result.DeletedItems.Length)
			' ExEnd:SynchronizeFolderItems
		End Sub
	End Class
End Namespace