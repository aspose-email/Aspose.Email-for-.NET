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

Public Class ParseMessageAndSave
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_POP3()
        Dim dstEmail As String = dataDir & Convert.ToString("first-message.eml")

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

        Try
            'Fetch the message by its sequence number
            Dim msg As MailMessage = client.FetchMessage(1)

            'Save the message using its subject as the file name
            msg.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultEml)

            client.Disconnect()
        Catch ex As Exception
            Console.WriteLine(Environment.NewLine + ex.Message)
        Finally
            client.Disconnect()
        End Try

        Console.WriteLine(Environment.NewLine + "Downloaded email using POP3. Message saved at " & dstEmail)
    End Sub
End Class
