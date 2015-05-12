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
		Friend Sub New(ByVal name As String, ByVal eMail As String, ByVal password As String, ByVal domain As String)
			Me.Name = name
			Me.EMail = eMail
			Me.Password = password
			Me.Domain = domain
		End Sub

		Public ReadOnly Name As String
		Public ReadOnly EMail As String
		Public ReadOnly Password As String
		Public ReadOnly Domain As String

		Public Shared Operator =(ByVal x As TestUser, ByVal y As TestUser) As Boolean
			If CObj(x) IsNot Nothing Then
				Return x.Equals(y)
			End If
			If CObj(y) IsNot Nothing Then
				Return y.Equals(x)
			End If
			Return True
		End Operator

		Public Shared Operator <>(ByVal x As TestUser, ByVal y As TestUser) As Boolean
			Return Not(x = y)
		End Operator

		Public Shared Widening Operator CType(ByVal user As TestUser) As String
			Return If(user Is Nothing, Nothing, user.Name)
		End Operator

		Public Overrides Function GetHashCode() As Integer
			Return ToString().GetHashCode()
		End Function

		Public Overrides Overloads Function Equals(ByVal obj As Object) As Boolean
			Return obj IsNot Nothing AndAlso TypeOf obj Is TestUser AndAlso Me.ToString().Equals(obj.ToString(), StringComparison.OrdinalIgnoreCase)
		End Function

		''' <summary>
		''' Returns a string that represents the current object.
		''' </summary>
		''' <returns>A string that represents the current object.</returns>
		Public Overrides Function ToString() As String
			Return If(String.IsNullOrEmpty(Domain), Name, String.Format("{0}/{1}", Domain, Name))
		End Function
	End Class
End Namespace