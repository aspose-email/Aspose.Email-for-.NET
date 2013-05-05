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
Imports System.Data
Imports System

Namespace PerformingMailMerge
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Create a new MailMessage instance
			Dim msg As New MailMessage()

			'Add subject and from address
			msg.Subject = "Hello, #FirstName#"
			msg.From = "sender@sender.com"

			'Add email address to send email
			msg.To.Add("asposetest123@gmail.com")

			'Add mesage field to HTML body
			msg.HtmlBody = "Your message here"
			msg.HtmlBody &= "Thank you for your interest in <STRONG>Aspose.Email</STRONG>."

			'Use GetSignment as the template routine, which will provide the same signature
			msg.HtmlBody &= "<br><br>Have fun with it.<br><br>#GetSignature()#"

			'Create a new TemplateEngine with the MSG message.
			Dim engine As New TemplateEngine(msg)

			' Register GetSignature routine. It will be used in MSG.
			engine.RegisterRoutine("GetSignature", New TemplateRoutine(AddressOf GetSignature))

			'Create an instance of DataTable
			'Fill a DataTable as data source
			Dim dt As New DataTable()
			dt.Columns.Add("Receipt", GetType(String))
			dt.Columns.Add("FirstName", GetType(String))
			dt.Columns.Add("LastName", GetType(String))

			'Create an instance of DataRow
			Dim dr As DataRow

			dr = dt.NewRow()
			dr("Receipt") = "abc<asposetest123@gmail.com>"
			dr("FirstName") = "a"
			dr("LastName") = "bc"
			dt.Rows.Add(dr)

			dr = dt.NewRow()
			dr("Receipt") = "Gunagzhou Team<asposetest123@gmail.com>"
			dr("FirstName") = "Guangzhou"
			dr("LastName") = "Team"
			dt.Rows.Add(dr)

			dr = dt.NewRow()
			dr("Receipt") = "Kyle.Huang<kyle@somedomain.com>"
			dr("FirstName") = "Kyle"
			dr("LastName") = "Huang"
			dt.Rows.Add(dr)

			Dim messages As MailMessageCollection
			Try
				'Create messages from the message and datasource.
				messages = engine.Instantiate(dt)

				'Create an instance of SmtpClient and specify server, port, username and password
				Dim client As SmtpClient = GetSmtpClient()

				'Send messages in bulk
				client.BulkSend(messages)
			Catch ex As MailException
				System.Diagnostics.Debug.WriteLine(ex.ToString())

			Catch ex As SmtpException
				System.Diagnostics.Debug.WriteLine(ex.ToString())
			End Try



		End Sub
		'Template routine to provide signature
		Private Shared Function GetSignature(ByVal args() As Object) As Object
			Return "Aspose.Email<br>Aspose Guangzhou Team<br>Aspose Ltd.<br>" & DateTime.Now.ToShortDateString()
		End Function
		Private Shared Function GetSmtpClient() As SmtpClient
			Dim client As New SmtpClient("smtp.gmail.com", 587, "asposetest123@gmail.com", "F123456f")
			client.SecurityMode = SmtpSslSecurityMode.Explicit
			client.EnableSsl = True

			Return client
		End Function
	End Class
End Namespace