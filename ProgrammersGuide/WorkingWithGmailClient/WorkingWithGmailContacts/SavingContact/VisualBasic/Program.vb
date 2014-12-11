'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Email
Imports Aspose.Email.Tests
Imports Aspose.Email.Google
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

Namespace SavingContactExample
	Public Class Program
		Public Shared Sub Main()
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")
			Dim username As String = "aspose.examples"
			Dim email As String = "aspose.examples@gmail.com"
			Dim password As String = "aspose123"
			Dim clientId As String = "491198589534-c08fdgnpu3ausn9pbdjjjh9r2s4vlla0.apps.googleusercontent.com"
			Dim clientSecret As String = "Km9G7a6PivZv4dDxZ3-ZVPk2"
			Dim refresh_token As String = "1/UlvQ1pvWN5DQLgHd8LmcNchPs50s7E96sTnjdwHKVS8"


			'The refresh_token is to be used in below.
			Dim user As New GoogleTestUser(username, email, password, clientId, clientSecret, refresh_token) 'refresh token

			Using client As IGmailClient = Aspose.Email.Google.GmailClient.GetInstance(user.ClientId, user.ClientSecret, user.RefreshToken)

				'Contact contact = new Contact();
				Dim contacts() As Contact = client.GetAllContacts()
				Dim contact As Contact = contacts(0)
				contact.Save("contact.msg", ContactSaveFormat.Msg)
				contact.Save("contact.vcf", ContactSaveFormat.VCard)


			End Using
		End Sub
	End Class

End Namespace