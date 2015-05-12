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

'INSTANT VB NOTE: The parameter clientId was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
'INSTANT VB NOTE: The parameter clientSecret was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
		Public Sub New(ByVal name As String, ByVal eMail As String, ByVal password As String, ByVal clientId_Renamed As String, ByVal clientSecret_Renamed As String)
			Me.New(name, eMail, password, clientId_Renamed, clientSecret_Renamed, Nothing)
		End Sub

'INSTANT VB NOTE: The parameter clientId was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
'INSTANT VB NOTE: The parameter clientSecret was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
'INSTANT VB NOTE: The parameter refreshToken was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
		Public Sub New(ByVal name As String, ByVal eMail As String, ByVal password As String, ByVal clientId_Renamed As String, ByVal clientSecret_Renamed As String, ByVal refreshToken_Renamed As String)
			MyBase.New(name, eMail, password, "gmail.com")
			ClientId = clientId_Renamed
			ClientSecret = clientSecret_Renamed
			RefreshToken = refreshToken_Renamed
		End Sub

		Public ReadOnly ClientId As String
		Public ReadOnly ClientSecret As String
		Public ReadOnly RefreshToken As String
	End Class
End Namespace