Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
	Class ExchangeFoldersBackupToPST
		Public Shared Sub Run()
			' ExStart:ExchangeFoldersBackupToPST
			Dim dataDir As String = RunExamples.GetDataDir_Exchange()
			' Create instance of IEWSClient class by providing credentials
			Const  mailboxUri As String = "https://ews.domain.com/ews/Exchange.asmx"
			Const  domain As String = ""
			Const  username As String = "username"
			Const  password As String = "password"
			Dim credential As New NetworkCredential(username, password, domain)
			Dim client As IEWSClient = EWSClient.GetEWSClient(mailboxUri, credential)

			' Get Exchange mailbox info of other email account
			Dim mailboxInfo As ExchangeMailboxInfo = client.GetMailboxInfo()
			Dim info As ExchangeFolderInfo = client.GetFolderInfo(mailboxInfo.InboxUri)
			Dim fc As New ExchangeFolderInfoCollection()
			fc.Add(info)
			client.Backup(fc, dataDir & Convert.ToString("Backup_out.pst"), Aspose.Email.Outlook.Pst.BackupOptions.None)
			' ExEnd:ExchangeFoldersBackupToPST
		End Sub
	End Class
End Namespace