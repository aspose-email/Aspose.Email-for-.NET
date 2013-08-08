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

Namespace CustomizingEmailHeaders
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
			Directory.CreateDirectory(dataDir)

			'Create an instance MailMessage class
			Dim msg As New MailMessage()

			'Specify ReplyTo
			msg.ReplyToList.Add("reply@reply.com")

			'From field
			msg.From = "sender@sender.com"

			'To field
			msg.To.Add("receiver1@receiver.com")

			'Adding Cc and Bcc Addresses
			msg.CC.Add("receiver2@receiver.com")
			msg.Bcc.Add("receiver3@receiver.com")

			'Message subject
			msg.Subject = "test mail"

			'Specify Date
			msg.Date = New System.DateTime(2006, 3, 6)

			'Specify XMailer
			msg.XMailer = "Aspose.Email"

			'Specify Secret Header
			msg.Headers.Add("secret-header", "mystery")

			'Save message to disc
			msg.Save(dataDir & "MsgHeaders.msg", MessageFormat.Msg)
		End Sub
	End Class
End Namespace