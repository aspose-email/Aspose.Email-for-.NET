'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Email
Imports System
Imports Aspose.Email.Imap

Namespace GettingFoldersInformation
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Create an instance of the ImapClient class
			Dim client As New ImapClient()

			'Specify host, username and password for your client
			client.Host = "imap.gmail.com"

			' Set username
			client.Username = "asposetest123@gmail.com"

			' Set password
			client.Password = "F123456f"

			' Set the port to 993. This is the SSL port of IMAP server
			client.Port = 993

			' Enable SSL
			client.EnableSsl = True

			Try
				'Connect to the remote server.
				client.Connect()

				'Log in to the remote server.
				client.Login()

				' Get all folders in the currently subscribed folder
				Dim folderInfoColl As Aspose.Email.Imap.ImapFolderInfoCollection = client.ListFolders()

				' Iterate through the collection to get folder info one by one
				For Each folderInfo As Aspose.Email.Imap.ImapFolderInfo In folderInfoColl
					' Folder name
					Console.WriteLine("Folder name is " & folderInfo.Name)
					Dim folderExtInfo As ImapFolderInfo = client.ListFolder(folderInfo.Name)
					' New messages in the folder
					Console.WriteLine("New message count: " & folderExtInfo.NewMessageCount)
					' Check whether its readonly
					Console.WriteLine("Is it readonly? " & folderExtInfo.ReadOnly)
					' Total number of messages
					Console.WriteLine("Total number of messages " & folderExtInfo.TotalMessageCount)
				Next folderInfo

				'Disconnect to the remote IMAP server
				client.Disconnect()

			Catch ex As Exception
				System.Console.Write(ex.ToString())
			End Try
		End Sub
	End Class
End Namespace