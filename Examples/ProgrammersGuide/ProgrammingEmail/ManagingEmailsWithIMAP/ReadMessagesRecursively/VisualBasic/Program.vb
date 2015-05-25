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
Imports System
Imports Aspose.Email.Mail

Namespace ReadMessagesRecursively
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)

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
                System.Console.WriteLine("Logged in to the IMAP server")

				' The root folder (which will be created on disk) consists of host and username
				Dim rootFolder As String = client.Host & "-" & client.Username
				' Create the root folder
				Directory.CreateDirectory(rootFolder)

				' List all the folders from IMAP server
				Dim folderInfoCollection As ImapFolderInfoCollection = client.ListFolders()
				For Each folderInfo As ImapFolderInfo In folderInfoCollection
					' Call the recursive method to read messages and get sub-folders
					ListMessagesInFolder(folderInfo, rootFolder,client)
				Next folderInfo


				'Disconnect to the remote IMAP server
				client.Disconnect()

				System.Console.WriteLine("Disconnected from the IMAP server")
			Catch ex As Exception
				System.Console.Write(ex.ToString())
			End Try
		End Sub
		''' <summary>
		''' Recursive method to get messages from folders and sub-folders
		''' </summary>
		''' <param name="folderInfo"></param>
		''' <param name="rootFolder"></param>
		Private Shared Sub ListMessagesInFolder(ByVal folderInfo As ImapFolderInfo, ByVal rootFolder As String, ByVal client As ImapClient)
			' Create the folder in disk (same name as on IMAP server)
            Dim currentFolder As String = Path.Combine(Path.GetFullPath("../../../Data/"), folderInfo.Name)
			Directory.CreateDirectory(currentFolder)

			' Read the messages from the current folder, if it is selectable
			If folderInfo.Selectable = True Then
				' Send status command to get folder info
				Dim folderInfoStatus As ImapFolderInfo = client.ListFolder(folderInfo.Name)
				Console.WriteLine(folderInfoStatus.Name & " folder selected. New messages: " & folderInfoStatus.NewMessageCount & ", Total messages: " & folderInfoStatus.TotalMessageCount)

				' Select the current folder
				client.SelectFolder(folderInfo.Name)
				' List messages
				Dim msgInfoColl As ImapMessageInfoCollection = client.ListMessages()
				Console.WriteLine("Listing messages....")
				For Each msgInfo As ImapMessageInfo In msgInfoColl
					' Get subject and other properties of the message
					Console.WriteLine("Subject: " & msgInfo.Subject)
					Console.WriteLine("Read: " & msgInfo.IsRead & ", Recent: " & msgInfo.Recent & ", Answered: " & msgInfo.Answered)

					' Get rid of characters like ? and :, which should not be included in a file name
					Dim fileName As String = msgInfo.Subject.Replace(":", " ").Replace("?", " ")

					' Save the message in MSG format
					Dim msg As MailMessage = client.FetchMessage(msgInfo.SequenceNumber)
                    msg.Save(currentFolder & "\" & fileName & "-" & msgInfo.SequenceNumber & ".msg", Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode)
				Next msgInfo
				Console.WriteLine("============================" & Constants.vbLf)
			Else
				Console.WriteLine(folderInfo.Name & " is not selectable.")
			End If

			Try
				' If this folder has sub-folders, call this method recursively to get messages
				Dim folderInfoCollection As ImapFolderInfoCollection = client.ListFolders(folderInfo.Name)
				For Each subfolderInfo As ImapFolderInfo In folderInfoCollection
					ListMessagesInFolder(subfolderInfo, rootFolder,client)
				Next subfolderInfo
			Catch e1 As Exception
			End Try
		End Sub

	End Class
End Namespace