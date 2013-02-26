'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Email
Imports Aspose.Email.Outlook.Pst

Namespace NewPSTAddSubfolders
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create new PST
			Dim pst As PersonalStorage = PersonalStorage.Create(dataDir & "PersonalStorage.pst", FileFormatVersion.Unicode)

			' Add new folder "Inbox"
			pst.RootFolder.AddSubFolder("Inbox")

			' Display Status.
			System.Console.WriteLine("New PST file created and addition of folder done successfully in it.")
		End Sub
	End Class
End Namespace