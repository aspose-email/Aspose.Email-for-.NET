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

Namespace ConnectingToSSLEnabledSMTPServer
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim client As New Aspose.Email.Mail.SmtpClient("smtp.gmail.com")

			' Set username
			client.Username = "asposetest123@gmail.com"

			' Set password
			client.Password = "F123456f"

			' Set the port to 587. This is the SSL port of SMTP server
			client.Port = 587

			' Set the security mode to explicit
			client.SecurityMode = SmtpSslSecurityMode.Explicit

			' Enable SSL
			client.EnableSsl = True

			'Declare msg as MailMessage instance
			Dim msg As New MailMessage()

			'use MailMessage properties like specify sender, recipient and message
			msg.To = "newcustomeronnet@gmail.com"
			msg.From = "newcustomeronnet@gmail.com"
			msg.Subject = "Test Email"
			msg.Body = "Hello World!"

			Try
				'Client will send this message
				client.Send(msg)
				'Show only if message sent successfully
				System.Console.WriteLine("Message sent")

			Catch ex As System.Exception
				System.Diagnostics.Trace.WriteLine(ex.ToString())
			End Try

		End Sub
	End Class
End Namespace