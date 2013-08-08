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
Imports Aspose.Email.Pop3
Imports System

Namespace ConnectingToSSLEnabledPOP3Server
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

			'Set Security Mode
			client.SecurityMode = Pop3SslSecurityMode.Implicit

			'Connect and login to a POP3 server
			Try
				client.Connect()
				Console.WriteLine("==================")
				Console.WriteLine("connected ")
				Console.WriteLine("==================")
				client.Login()
				Console.WriteLine("==================")
				Console.WriteLine("logged in ")
				Console.WriteLine("==================")
			Catch ex As Pop3Exception
				Console.Write(ex.ToString())
			End Try


		End Sub
	End Class
End Namespace