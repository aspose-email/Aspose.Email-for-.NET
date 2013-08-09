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
Imports System.Configuration

Namespace LoadSmtpConfigFile
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Declare msg as MailMessage instance
			Dim msg As New MailMessage()

			'use MailMessage properties like specify sender, recipient and message
			msg.To = "asposetest123@gmail.com"
			msg.From = "aspose2@gmail.com"
			msg.Subject = "Test Email"
			msg.Body = "Hello World!"

			'Create an instance of SmtpClient class and load SMTP Authentication settings from Config file
			Dim client As New SmtpClient(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None))
			client.EnableSsl = True

			Try
				'Client.Send will send this message
				client.Send(msg)
                'Message sent successfully
				System.Console.WriteLine("Message sent")

			Catch ex As System.Exception
				System.Diagnostics.Trace.WriteLine(ex.ToString())
			End Try
		End Sub
	End Class
End Namespace