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
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

Namespace CreateSaveOutlookFiles
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
			Directory.CreateDirectory(dataDir)

			' Create an instance of MailMessage class
			Dim mailMsg As New MailMessage()

			' Set FROM field of the message
			mailMsg.From = "from@domain.com"

			' Set TO field of the message
			mailMsg.To.Add("to@domain.com")

			' Set SUBJECT of the message
			mailMsg.Subject = "creating an outlook message file"

			' Set BODY of the message
			mailMsg.Body = "This message is created by Aspose.Email"

			' Create an instance of MapiMessage class and pass MailMessage as argument
			Dim outlookMsg As MapiMessage = MapiMessage.FromMailMessage(mailMsg)

			' Save the message (msg) file
			outlookMsg.Save(dataDir & "message.msg")

			' Display Status.
			System.Console.WriteLine("MSG file created and saved successfully.")
		End Sub
	End Class
End Namespace