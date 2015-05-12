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
Imports Aspose.Email.Outlook

Namespace DraftAppointmentRequest
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
			Directory.CreateDirectory(dataDir)

			Dim sender As String = "test@gmail.com"
			Dim recipient As String = "test@email.com"

			Dim message As New MailMessage(sender, recipient, String.Empty, String.Empty)

			Dim app As New Appointment(String.Empty, DateTime.Now, DateTime.Now, sender, recipient)
			app.Method = AppointmentMethodType.Publish

			message.AddAlternateView(app.RequestApointment())

			Dim msg As MapiMessage = MapiMessage.FromMailMessage(message)

			' Save the appointment as draft.
			msg.Save(dataDir & "draftAppointment.msg")
		End Sub
	End Class
End Namespace