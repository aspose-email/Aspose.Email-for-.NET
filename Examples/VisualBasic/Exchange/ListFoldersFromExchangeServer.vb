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
	Class ListFoldersFromExchangeServer
		' ExStart:ListFoldersFromExchangeServer
		Public Shared Sub Run()
			Try
				Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
				Console.WriteLine("Downloading all messages from Inbox....")

				Dim mailboxInfo As ExchangeMailboxInfo = client.GetMailboxInfo()
				Console.WriteLine("Mailbox URI: " + mailboxInfo.MailboxUri)
				Dim rootUri As String = client.GetMailboxInfo().RootUri
				' List all the folders from Exchange server
				Dim folderInfoCollection As ExchangeFolderInfoCollection = client.ListSubFolders(rootUri)
				For Each folderInfo As ExchangeFolderInfo In folderInfoCollection
					' Call the recursive method to read messages and get sub-folders
					ListSubFolders(client, folderInfo)
				Next

				Console.WriteLine("All messages downloaded.")
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			End Try
		End Sub
		Private Shared Sub ListSubFolders(client As IEWSClient, folderInfo As ExchangeFolderInfo)
			' Create the folder in disk (same name as on IMAP server)
			Console.WriteLine(folderInfo.DisplayName)
			Try
				' If this folder has sub-folders, call this method recursively to get messages
				Dim folderInfoCollection As ExchangeFolderInfoCollection = client.ListSubFolders(folderInfo.Uri)
				For Each subfolderInfo As ExchangeFolderInfo In folderInfoCollection
					ListSubFolders(client, subfolderInfo)
				Next
			Catch generatedExceptionName As Exception
			End Try
		End Sub
		' ExEnd:ListFoldersFromExchangeServer
	End Class
End Namespace