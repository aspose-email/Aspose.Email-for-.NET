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
Imports Aspose.Email.Pop3
Imports System

Namespace GettingMailboxInfo
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Create an instance of the Pop3Client class
			Dim client As New Pop3Client()

			'Specify host, username and password for your client
			client.Host = "pop.gmail.com"

			' Set username
			client.Username = "asposetest123@gmail.com"

			' Set password
			client.Password = "F123456f"

			' Set the port to 995. This is the SSL port of POP3 server
			client.Port = 995

			' Enable SSL
			client.EnableSsl = True

			' Connect to a POP3 server
			client.Connect(True)

			' Get the size of the mailbox
			Dim nSize As Long = client.GetMailboxSize()

			Console.WriteLine("Mailbox size is " & nSize & " bytes.")
			' Get mailbox info

			Dim info As Aspose.Email.Pop3.Pop3MailboxInfo = client.GetMailboxInfo()

			' Get the number of messages in the mailbox
			Dim nMessageCount As Integer = info.MessageCount

			Console.WriteLine("Number of messages in mailbox are " & nMessageCount)

			' Get occupied size
			Dim nOccupiedSize As Long = info.OccupiedSize

			Console.WriteLine("Occupied size is " & nOccupiedSize)

		End Sub
	End Class
End Namespace