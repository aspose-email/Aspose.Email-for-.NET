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
	Class DownloadMessagesFromPublicFolders
		' ExStart:DownloadMessagesFromPublicFolders
		Public Shared dataDir As String = RunExamples.GetDataDir_Exchange()
		Public Shared mailboxUri As String = "https://exchange/ews/exchange.asmx"
		Public Shared username As String = "administrator"
		Public Shared password As String = "pwd"
		Public Shared domain As String = "ex2013.local"

		Public Shared Sub Run()
			Try
				ReadPublicFolders()
			Catch ex As Exception
				Console.WriteLine(ex.Message)
			End Try
		End Sub

		Private Shared Sub ReadPublicFolders()
			Dim credential As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credential)

			Dim folders As ExchangeFolderInfoCollection = client.ListPublicFolders()
			For Each publicFolder As ExchangeFolderInfo In folders
				Console.WriteLine("Name: " + publicFolder.DisplayName)
				Console.WriteLine("Subfolders count: " + publicFolder.ChildFolderCount)

				ListMessagesFromSubFolder(publicFolder, client)
			Next
		End Sub

		Private Shared Sub ListMessagesFromSubFolder(publicFolder As ExchangeFolderInfo, client As IEWSClient)
			Console.WriteLine("Folder Name: " + publicFolder.DisplayName)
			Dim msgInfoCollection As ExchangeMessageInfoCollection = client.ListMessagesFromPublicFolder(publicFolder)
			For Each messageInfo As ExchangeMessageInfo In msgInfoCollection
				Dim msg As MailMessage = client.FetchMessage(messageInfo.UniqueUri)
				Console.WriteLine(msg.Subject)
				msg.Save(dataDir + msg.Subject & Convert.ToString(".msg"), SaveOptions.DefaultMsgUnicode)
			Next

			' Call this method recursively for any subfolders
			If publicFolder.ChildFolderCount > 0 Then
				Dim subfolders As ExchangeFolderInfoCollection = client.ListSubFolders(publicFolder)
				For Each subfolder As ExchangeFolderInfo In subfolders
					ListMessagesFromSubFolder(subfolder, client)
				Next
			End If
		End Sub
		' ExEnd:DownloadMessagesFromPublicFolders
	End Class
End Namespace