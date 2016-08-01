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
	Class CreateFoldersOnExchangeServerMailbox
		Public Shared Sub Run()
			' ExStart:CreateFoldersOnExchangeServerMailbox

			' Create instance of EWSClient class by giving credentials
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
			Dim inbox As String = client.MailboxInfo.InboxUri
			Dim folderName1 As String = "EMAILNET-35054"
			Dim subFolderName0 As String = "2015"
			Dim folderName2 As String = Convert.ToString(folderName1 & Convert.ToString("/")) & subFolderName0
			Dim folderName3 As String = folderName1 & Convert.ToString(" / 2015")
			Dim rootFolderInfo As ExchangeFolderInfo = Nothing
			Dim folderInfo As ExchangeFolderInfo = Nothing

			Try
				client.UseSlashAsFolderSeparator = True
				client.CreateFolder(client.MailboxInfo.InboxUri, folderName1)
				client.CreateFolder(client.MailboxInfo.InboxUri, folderName2)
			Finally
				If client.FolderExists(inbox, folderName1, rootFolderInfo) Then
					If client.FolderExists(inbox, folderName2, folderInfo) Then
						client.DeleteFolder(folderInfo.Uri, True)
					End If
					client.DeleteFolder(rootFolderInfo.Uri, True)
				End If
			End Try
			' ExEnd:CreateFoldersOnExchangeServerMailbox
		End Sub
	End Class
End Namespace