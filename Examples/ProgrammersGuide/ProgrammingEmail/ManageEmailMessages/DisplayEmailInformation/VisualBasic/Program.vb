'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Email
Imports Aspose.Email.Mail

Namespace DisplayEmailInformation
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
            ' The path to the documents directory.
            Dim dataDir As String = Path.GetFullPath("../../../Data/")

            Dim message As MailMessage

            'Create MailMessage instance by loading an Eml file
            message = MailMessage.Load(dataDir & Convert.ToString("test.eml"), MailMessageLoadOptions.DefaultEml)

            System.Console.Write("From:")

            'Gets the sender info
            System.Console.WriteLine(message.From)
            System.Console.Write("To:")

            'Gets the recipient info
            System.Console.WriteLine(message.[To])
            System.Console.Write("Subject:")

            'Gets the subject
            System.Console.WriteLine(message.Subject)
            System.Console.WriteLine("HtmlBody:")

            'Gets the htmlbody 
            System.Console.WriteLine(message.HtmlBody)
            System.Console.WriteLine("TextBody")

            'Gets the textbody
            System.Console.WriteLine(message.Body)
        End Sub
    End Class
End Namespace