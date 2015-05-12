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
Imports Aspose.Email.Outlook
Imports System

Namespace RecipientInformation
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Load message file
			Dim msg As MapiMessage = MapiMessage.FromFile(dataDir & "Message.msg")

			'Enumerate the recipients
			For Each recip As MapiRecipient In msg.Recipients
				Select Case recip.RecipientType ' What's the type?
					Case MapiRecipientType.MAPI_TO
						Console.WriteLine("RecipientType:TO")
					Case MapiRecipientType.MAPI_CC
						Console.WriteLine("RecipientType:CC")
					Case MapiRecipientType.MAPI_BCC
						Console.WriteLine("RecipientType:BCC")
				End Select

				'Get email address
				Console.WriteLine("Email Address: " & recip.EmailAddress)
				'Get display name
				Console.WriteLine("DisplayName: " & recip.DisplayName)
				'Get address type
				Console.WriteLine("AddressType: " & recip.AddressType)
			Next recip

		End Sub
	End Class
End Namespace