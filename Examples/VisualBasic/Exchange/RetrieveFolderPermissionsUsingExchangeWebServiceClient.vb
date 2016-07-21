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
	Class RetrieveFolderPermissionsUsingExchangeWebServiceClient
		Public Shared Sub Run()
			' ExStart:RetrieveFolderPermissionsUsingExchangeWebServiceClient
			Dim folderName As String = "DesiredFolderName"

			' Create instance of EWSClient class by giving credentials
			Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

			Dim folders As ExchangeFolderInfoCollection = client.ListPublicFolders()
			Dim permissions As New ExchangeFolderPermissionCollection()
			Dim publicFolder As ExchangeFolderInfo = Nothing
			Try
				For Each folderInfo As ExchangeFolderInfo In folders
					If folderInfo.DisplayName.Equals(folderName) Then
						publicFolder = folderInfo
					End If
				Next

				If publicFolder Is Nothing Then
					Console.WriteLine("public folder was not created in the root public folder")
				End If

				Dim folderPermissionCol As ExchangePermissionCollection = client.GetFolderPermissions(publicFolder.Uri)
				For Each perm As ExchangeBasePermission In folderPermissionCol
					Dim permission As ExchangeFolderPermission = TryCast(perm, ExchangeFolderPermission)
					If permission Is Nothing Then
						Console.WriteLine("Permission is null.")
					Else
						Console.WriteLine("User's primary smtp address: {0}", permission.UserInfo.PrimarySmtpAddress)
						Console.WriteLine("User can create Items: {0}", permission.CanCreateItems.ToString())
						Console.WriteLine("User can delete Items: {0}", permission.DeleteItems.ToString())
						Console.WriteLine("Is Folder Visible: {0}", permission.IsFolderVisible.ToString())
						Console.WriteLine("Is User owner of this folder: {0}", permission.IsFolderOwner.ToString())
						Console.WriteLine("User can read items: {0}", permission.ReadItems.ToString())
					End If
				Next
				Dim mailboxInfo As ExchangeMailboxInfo = client.GetMailboxInfo()
				'Get the Permissions for the Contacts and Calendar Folder
				Dim contactsPermissionCol As ExchangePermissionCollection = client.GetFolderPermissions(mailboxInfo.ContactsUri)
				Dim calendarPermissionCol As ExchangePermissionCollection = client.GetFolderPermissions(mailboxInfo.CalendarUri)
					'Do the needfull
			Finally
			End Try
			' ExEnd:RetrieveFolderPermissionsUsingExchangeWebServiceClient
		End Sub
	End Class
End Namespace