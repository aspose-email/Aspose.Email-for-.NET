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
Imports Aspose.Email.Mime
Imports System
Imports Aspose.Email.Pop3

Namespace RetrievingEmailHeaders
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Create an instance of the Pop3Client class
			Dim client As New Pop3Client()

			'Specify host, username and password for your client
			client.Host = "pop.gmail.com"

			' Set username
			client.Username = "asposetest123@gmail.com"

			' Set password
			client.Password = "F123456f"

			' Set the port to 995. This is the SSL port of POP3 server
			client.Port = 995

			' Enable SSL
			client.EnableSsl = True

			Try
				client.Connect(True)

				Dim headers As HeaderCollection = client.GetMessageHeaders(1)

				For i As Integer = 0 To headers.Count - 1
					' Display key and value in the header collection
					Console.Write(headers.Keys(i))
					Console.Write(" : ")
					Console.WriteLine(headers.Get(i))
				Next i

			Catch ex As Exception
				Console.WriteLine(ex.Message)
			Finally
				client.Disconnect()
			End Try
		End Sub
	End Class
End Namespace