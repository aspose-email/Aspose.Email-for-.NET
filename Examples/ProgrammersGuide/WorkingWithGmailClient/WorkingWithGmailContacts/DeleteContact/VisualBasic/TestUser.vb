Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Net


Namespace Aspose.Email.Tests.TestUtils
	Friend Enum GrantTypes
		authorization_code
		refresh_token
	End Enum
End Namespace

Namespace Aspose.Email.Tests
	Public Class TestUser
'INSTANT VB NOTE: The parameter name was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
'INSTANT VB NOTE: The parameter eMail was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
'INSTANT VB NOTE: The parameter password was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
'INSTANT VB NOTE: The parameter domain was renamed since Visual Basic will not uniquely identify class members when parameters have the same name:
		Friend Sub New(ByVal name_Renamed As String, ByVal eMail_Renamed As String, ByVal password_Renamed As String, ByVal domain_Renamed As String)
			Name = name_Renamed
			EMail = eMail_Renamed
			Password = password_Renamed
			Domain = domain_Renamed
		End Sub

		Public ReadOnly Name As String
		Public ReadOnly EMail As String
		Public ReadOnly Password As String
		Public ReadOnly Domain As String

		Public Shared Operator =(ByVal x As TestUser, ByVal y As TestUser) As Boolean
			If Not CObj(x) Is Nothing Then
				Return x.Equals(y)
			End If
			If Not CObj(y) Is Nothing Then
				Return y.Equals(x)
			End If
			Return True
		End Operator

		Public Shared Operator <>(ByVal x As TestUser, ByVal y As TestUser) As Boolean
			Return Not(x Is y)
		End Operator

		Public Shared Widening Operator CType(ByVal user As TestUser) As String
			If user Is Nothing Then
				Return Nothing
			Else
				Return user.Name
			End If
		End Operator

		Public Overrides Function GetHashCode() As Integer
			Return ToString().GetHashCode()
		End Function

		Public Overrides Function Equals(ByVal obj As Object) As Boolean
			Return Not obj Is Nothing AndAlso TypeOf obj Is TestUser AndAlso Me.ToString().Equals(obj.ToString(), StringComparison.OrdinalIgnoreCase)
		End Function

		''' <summary>
		''' Returns a string that represents the current object.
		''' </summary>
		''' <returns>A string that represents the current object.</returns>
		Public Overrides Function ToString() As String
			If String.IsNullOrEmpty(Domain) Then
				Return Name
			Else
				Return String.Format("{0}/{1}", Domain, Name)
			End If
		End Function
	End Class
End Namespace