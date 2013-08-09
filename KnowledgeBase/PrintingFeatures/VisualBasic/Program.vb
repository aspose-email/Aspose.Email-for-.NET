'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO

Imports Aspose.Email
Imports Aspose.Email.Mail
Imports Aspose.Email.Printing

Namespace PrintingFeatures
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Directory.CreateDirectory(dataDir)

			'Declare message as an MailMessage instance
			Dim message As New MailMessage()

			'Sender's address
			message.From = "user1@domain.com"

			'Recipient's address
			message.To = "user2@domain.com"

			'Subject of the email
			message.Subject = "My First Mail"

			'Set message date to current date and time
			message.Date = DateTime.Now

			'Body of the email message
			message.TextBody = "Text is the Mail Message"

			'Body of the email message
			message.IsBodyHtml = True
			message.HtmlBody = "<html><body><h1>Hello this is html body in Heading 1 format </h1></body></html>"

			'Instantiate an instance of MailPrinter
			Dim printer As New Aspose.Email.Printing.MailPrinter()

			'Set the MessageFormattingFlags to none to display only the message body
			printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.None

			'Set MessageFormattingFlags to MailInfo to display the message headers and body
			printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.MailInfo

			'Just for testing, get the default property values
			Dim width As Double = printer.PageWidth
			Dim height As Double = printer.PageHeight

			'Set page layout for printing
			printer.PageUnit = Aspose.Email.Printing.PrinterUnit.Cm
			printer.MarginTop = 2
			printer.MarginBottom = 2
			printer.MarginLeft = 2
			printer.MarginRight = 2
			printer.PageWidth = 8
			printer.PageHeight = 20

			'Print the email to an XPS file
			printer.Print(message, dataDir & "test.xps", Aspose.Email.Printing.PrintFormat.XPS)

			' Auto-Fit a TIFF
			printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.AutoFitWidth

			'Print the email to a TIFF file
			printer.Print(message, dataDir & "test.tiff", Aspose.Email.Printing.PrintFormat.Tiff)

			' Display Status.
			System.Console.WriteLine("Printing to XPS and TIFF completed successfully.")
		End Sub
	End Class
End Namespace