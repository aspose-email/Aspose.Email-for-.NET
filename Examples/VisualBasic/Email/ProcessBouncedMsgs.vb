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

Namespace Aspose.Email.Examples.VisualBasic.Email
    Public Class ProcessBouncedMsgs
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Email()
            Dim dstEmail As String = dataDir & Convert.ToString("test.eml")

            Dim fileName As String = dstEmail
            Dim mail As MailMessage = MailMessage.Load(fileName)
            Dim result As BounceResult = mail.CheckBounced()
            Console.WriteLine(fileName)
            Console.WriteLine("IsBounced : " + result.IsBounced.ToString())
            Console.WriteLine("Action : " + result.Action.ToString())
            Console.WriteLine("Recipient : " + result.Recipient)

            Console.WriteLine(Environment.NewLine + "Bounce information displayed successfully.")
        End Sub
    End Class
End Namespace