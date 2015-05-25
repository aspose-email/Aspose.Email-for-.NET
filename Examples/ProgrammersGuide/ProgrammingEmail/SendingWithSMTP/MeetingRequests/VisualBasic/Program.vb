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

Namespace MeetingRequests
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Create an instance of the MailMessage class
			Dim msg As New MailMessage()

			'set the sender
			msg.From = "newcustomeronnet@gmail.com"

			'Set the recipient, who will receive the meeting request.
			'Basically, the recipient is the same as the meeting attendees.
			msg.To = "person1@domain.com, person2@domain.com, person3@domain.com, asposetest123@gmail.com"

			'create Appointment instance
			Dim app As New Appointment("Room 112", New DateTime(2006, 7, 17, 13, 0, 0), New DateTime(2006, 7, 17, 14, 0, 0), msg.From, msg.To)

			app.Summary = "Release Meetting"
			app.Description = "Discuss for the next release"

			'add appointment to the message
			msg.AddAlternateView(app.RequestApointment())

			'Create an instance of SmtpClient class
			Dim client As SmtpClient = GetSmtpClient()

			Try
				'Client.Send will send this message
				client.Send(msg)
				'Show Message Sent�E only if message sent successfully
				Console.WriteLine("Message sent")

			Catch ex As Exception
				System.Diagnostics.Trace.WriteLine(ex.ToString())
			End Try

		End Sub
		Private Shared Function GetSmtpClient() As SmtpClient
			Dim client As New SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f")
			client.SecurityOptions = SecurityOptions.Auto

			Return client
		End Function
	End Class
End Namespace