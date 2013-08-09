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
Imports System

Namespace SettingTextBody
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Declare msg as MailMessage instance
			Dim msg As New MailMessage()

			'use MailMessage properties like specify sender, recipient and message
			'use MailMessage properties like specify sender, recipient and message
			msg.From = "newcustomeronnet@gmail.com"
			msg.To = "newcustomeronnet2@gmail.com"
			msg.Subject = "Test subject"
			msg.TextBody = "This is text body"


			Dim client As New SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f")

			client.SecurityMode = SmtpSslSecurityMode.Explicit

			client.EnableSsl = True

			Try
				'Client will send this message
				client.Send(msg)
				'Show only if message sent successfully
				Console.WriteLine("Message sent")

			Catch ex As Exception
				System.Diagnostics.Trace.WriteLine(ex.ToString())
			End Try

		End Sub
	End Class
End Namespace