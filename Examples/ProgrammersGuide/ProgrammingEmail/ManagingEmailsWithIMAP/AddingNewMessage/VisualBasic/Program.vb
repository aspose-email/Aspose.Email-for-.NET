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

Namespace AddingNewMessage
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create a message
			Dim msg As Aspose.Email.Mail.MailMessage
			msg = New Aspose.Email.Mail.MailMessage("user@domain1.com", "user@domain2.com", "subject", "message")

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
            client.SecurityOptions = SecurityOptions.Auto

			Try
                ' Subscribe to the Inbox folder
				client.SelectFolder(ImapFolderInfo.InBox)
				client.SubscribeFolder(client.CurrentFolder.Name)

				' Append the newly created message
				client.AppendMessage(client.CurrentFolder.Name, msg)

				System.Console.WriteLine("New Message Added Successfully")

				'Disconnect to the remote IMAP server
				client.Disconnect()

				System.Console.WriteLine("Disconnected from the IMAP server")
			Catch ex As Exception
				System.Console.Write(ex.ToString())
			End Try



		End Sub
	End Class
End Namespace