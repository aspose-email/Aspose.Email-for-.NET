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

Namespace Aspose.Email.Examples.VisualBasic.Exchange

    Public Class SendExchangeTask
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Exchange()
            Dim dstEmail As String = dataDir & Convert.ToString("Message.msg")

            ' Create instance of ExchangeClient class by giving credentials
            Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

            Dim loadOptions As New MailMessageLoadOptions()
            loadOptions.MessageFormat = MessageFormat.Msg
            loadOptions.FileCompatibilityMode = FileCompatibilityMode.PreserveTnefAttachments

            ' load task from .msg file
            Dim eml As MailMessage = MailMessage.Load(dstEmail, loadOptions)
            eml.From = "firstname.lastname@domain.com"
            eml.[To].Clear()
            eml.[To].Add(New Aspose.Email.Mail.MailAddress("firstname.lastname@domain.com"))

            client.Send(eml)

            Console.WriteLine(Environment.NewLine + "Task sent on Exchange Server successfully.")
        End Sub
    End Class
End Namespace