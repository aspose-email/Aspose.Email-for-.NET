Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text

Namespace Aspose.Email.Tests
	Public Class GoogleTestUser
		Inherits TestUser
		Public Sub New(ByVal name As String, ByVal eMail As String, ByVal password As String)
			Me.New(name, eMail, password, Nothing, Nothing, Nothing)
		End Sub

		Public Sub New(ByVal name As String, ByVal eMail As String, ByVal password As String, ByVal clientId As String, ByVal clientSecret As String)
			Me.New(name, eMail, password, clientId, clientSecret, Nothing)
		End Sub

		Public Sub New(ByVal name As String, ByVal eMail As String, ByVal password As String, ByVal clientId As String, ByVal clientSecret As String, ByVal refreshToken As String)
			MyBase.New(name, eMail, password, "gmail.com")
			Me.ClientId = clientId
			Me.ClientSecret = clientSecret
			Me.RefreshToken = refreshToken
		End Sub

		Public ReadOnly ClientId As String
		Public ReadOnly ClientSecret As String
		Public ReadOnly RefreshToken As String
	End Class
End Namespace