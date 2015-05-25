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

Namespace ExtractingEmailHeaders
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim message As MailMessage

			'Create MailMessage instance by loading an EML file
            message = MailMessage.Load(dataDir & "test.eml", MailMessageLoadOptions.DefaultEml)

			Console.WriteLine(Constants.vbLf + Constants.vbLf & "headers:" & Constants.vbLf + Constants.vbLf)

			'print out all the headers
			Dim index As Integer = 0
			For Each header As String In message.Headers
				Console.Write(header & " - ")
				Console.WriteLine(message.Headers.Get(index))
				index += 1
			Next header

			System.Console.Read()
		End Sub
	End Class
End Namespace