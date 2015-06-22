'''///////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Email. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'''///////////////////////////////////////////////////////////////////////
Imports Aspose.Email.Mail
Imports System.IO
Imports Aspose.Email.Mail.Bounce



Module Module1

    Sub Main()

        Dim dataDir As String = Path.GetFullPath("../../../Data/")

        Dim fileName As String = dataDir & Convert.ToString("failed1.eml")
        Dim mail As MailMessage = MailMessage.Load(fileName)
        Dim result As BounceResult = mail.CheckBounced()
        Console.WriteLine(fileName)
        Console.WriteLine("IsBounced : " + result.IsBounced)
        Console.WriteLine("Action : " + result.Action)
        Console.WriteLine("Recipient : " + result.Recipient)
        Console.WriteLine()

        fileName = dataDir & Convert.ToString("failedReport2.eml")
        mail = MailMessage.Load(fileName)
        result = mail.CheckBounced()
        Console.WriteLine(fileName)
        Console.WriteLine("IsBounced : " + result.IsBounced)
        Console.WriteLine("Action : " + result.Action)
        Console.WriteLine("Recipient : " + result.Recipient)
        Console.WriteLine("Reason : " + result.Reason)
        Console.WriteLine("Status : " + result.Status)
        Console.WriteLine("OriginalMessage ToAddress 1: " + result.OriginalMessage.[To](0).Address)
        Console.WriteLine("OriginalMessage ToAddress 2: " + result.OriginalMessage.[To](1).Address)
        Console.WriteLine()

        fileName = dataDir & Convert.ToString("delayed.eml")
        mail = MailMessage.Load(fileName)
        result = mail.CheckBounced()
        Console.WriteLine(fileName)
        Console.WriteLine("IsBounced : " + result.IsBounced)
        Console.WriteLine("Action : " + result.Action)
        Console.WriteLine("Recipient : " + result.Recipient)

        Console.WriteLine(vbLf & "Execution completed.")
        Console.ReadKey()
    End Sub

End Module
