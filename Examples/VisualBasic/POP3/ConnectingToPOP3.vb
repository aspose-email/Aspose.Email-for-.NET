' ///////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
' ///////////////////////////////////////////////////////////////////////

Imports System.IO
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook
Imports Aspose.Email.Pop3
Imports Aspose.Email
Imports Aspose.Email.Mime
Imports Aspose.Email.Imap
Imports System.Configuration
Imports System.Data
Imports Aspose.Email.Mail.Bounce
Imports Aspose.Email.Exchange
Imports Aspose.Email.Outlook.Pst

Public Class ConnectingToPOP3
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_POP3()
        Dim dstEmail As String = dataDir & Convert.ToString("1234.eml")

        'Create an instance of the Pop3Client class
        Dim client As New Pop3Client()

        'Specify host, username and password for your client
        client.Host = "pop.gmail.com"

        ' Set username
        client.Username = "your.username@gmail.com"

        ' Set password
        client.Password = "your.password"

        ' Set the port to 995. This is the SSL port of POP3 server
        client.Port = 995

        ' Enable SSL
        client.SecurityOptions = SecurityOptions.Auto

        Console.WriteLine(Environment.NewLine + "Connected to POP3 server.")
    End Sub
End Class
