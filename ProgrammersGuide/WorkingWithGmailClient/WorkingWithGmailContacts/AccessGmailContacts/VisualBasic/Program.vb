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
Imports Aspose.Email.Services.Google
Imports Aspose.Email.Google
Imports Aspose.Email.Mail
Imports System
Imports Aspose.Email.Tests

Namespace AccessGmailContactsExample
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
				Dim contacts() As Contact = client.GetAllContacts()
				For Each contact As Contact In contacts
					Console.WriteLine(contact.DisplayName & ", " & contact.EmailAddresses(0))
				Next contact

				'Fetch contacts from a specific group
				Dim groups As FeedEntryCollection = client.FetchAllGroups()
				Dim group As GmailContactGroup = Nothing
				For Each g As GmailContactGroup In groups
					Select Case g.Title
						Case "TestGroup"
							group = g
					End Select
				Next g

				'Retrieve contacts from the Group
				If group IsNot Nothing Then
					Dim contacts2() As Contact = client.GetContactsFromGroup(group)
					For Each con As Contact In contacts2
						Console.WriteLine(con.DisplayName & "," & con.EmailAddresses(0).ToString())
					Next con
				End If
			End Using
		End Sub
	End Class
End Namespace
