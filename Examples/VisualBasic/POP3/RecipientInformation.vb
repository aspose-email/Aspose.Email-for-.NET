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

Public Class RecipientInformation
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_POP3()
        Dim dstEmail As String = dataDir & Convert.ToString("message.msg")

        'Load message file
        Dim msg As MapiMessage = MapiMessage.FromFile(dstEmail)

        'Enumerate the recipients
        For Each recip As MapiRecipient In msg.Recipients
            Select Case recip.RecipientType
                ' What's the type?
                Case MapiRecipientType.MAPI_TO
                    Console.WriteLine("RecipientType:TO")
                    Exit Select
                Case MapiRecipientType.MAPI_CC
                    Console.WriteLine("RecipientType:CC")
                    Exit Select
                Case MapiRecipientType.MAPI_BCC
                    Console.WriteLine("RecipientType:BCC")
                    Exit Select
            End Select

            'Get email address
            Console.WriteLine("Email Address: " + recip.EmailAddress)
            'Get display name
            Console.WriteLine("DisplayName: " + recip.DisplayName)
            'Get address type
            Console.WriteLine("AddressType: " + recip.AddressType)
        Next

        Console.WriteLine(Environment.NewLine + "Displayed recipient information from MSG file " & dstEmail)
    End Sub
End Class
