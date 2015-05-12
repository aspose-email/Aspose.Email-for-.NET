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

Namespace MultipleRecipients
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Declare msg as MailMessage instance
			Dim message As New MailMessage()

			'use MailMessage properties like specify sender, recipient and message

			'Specify the recipients mail addresses
			message.To.Add("receiver1@receiver.com")
			message.To.Add("receiver2@receiver.com")
			message.To.Add("receiver3@receiver.com")
			message.To.Add("receiver4@receiver.com")

			message.CC.Add("CC1@receiver.com")
			message.CC.Add("CC2@receiver.com")

			message.Bcc.Add("Bcc1@receiver.com")
			message.Bcc.Add("Bcc2@receiver.com")

			message.From = "newcustomeronnet@gmail.com"
			message.Subject = "Test Email"
			message.Body = "Hello World!"


			'Create an instance of SmtpClient class
			Dim client As SmtpClient = GetSmtpClient()

			Try
				'Client will send this message
				client.Send(message)
				'Show only if message sent successfully
				Console.WriteLine("Message sent")

			Catch ex As Exception
				System.Diagnostics.Trace.WriteLine(ex.ToString())
			End Try

		End Sub

		Private Shared Function GetSmtpClient() As SmtpClient
			Dim client As New SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f")
			client.SecurityMode = SmtpSslSecurityMode.Explicit
			client.EnableSsl = True

			Return client
		End Function
	End Class
End Namespace