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
Imports Aspose.Email.Imap

Namespace MessagesFromIMAPServerToDisk
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
            Directory.CreateDirectory(dataDir)

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

			client.SecurityMode = ImapSslSecurityMode.Implicit

			' Enable SSL
			client.EnableSsl = True

			Try

				client.Connect()

				'Log in to the remote server.
				client.Login()

				' Select the inbox folder
				client.SelectFolder(ImapFolderInfo.InBox)

				' Get the message info collection
				Dim list As ImapMessageInfoCollection = client.ListMessages()

				' Download each message
				For i As Integer = 0 To list.Count - 1
					'Save the EML file locally
					client.SaveMessage(list(i).UniqueId, dataDir & list(i).UniqueId & ".eml")
				Next i

				'Disconnect to the remote IMAP server
				client.Disconnect()

				System.Console.WriteLine("Disconnected from the IMAP server")
			Catch ex As System.Exception
				System.Console.Write(ex.ToString())
			End Try



		End Sub
	End Class
End Namespace