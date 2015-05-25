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

Namespace ExportAsEML
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
            Directory.CreateDirectory(dataDir)

			'Create a new instance of MailMessage class
			Dim message As New MailMessage()

			' Set subject of the message
			message.Subject = "New message created by Aspose.Email for .NET"

			' Set Html body
			message.IsBodyHtml = True
			message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>"

			' Set sender information
			message.From = "from@domain.com"

			' Add TO recipients
			message.To.Add("to1@domain.com")
			message.To.Add("to2@domain.com")

			'Add CC recipients
			message.CC.Add("cc1@domain.com")
			message.CC.Add("cc2@domain.com")

			' Save message in EML, MSG and MHTML formats
            message.Save(dataDir & "Message.eml", Aspose.Email.Mail.SaveOptions.DefaultEml)
		End Sub
	End Class
End Namespace