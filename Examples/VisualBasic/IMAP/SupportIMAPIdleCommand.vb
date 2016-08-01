Imports System.Threading
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
	Class SupportIMAPIdleCommand
		Public Shared Sub Run()
			' ExStart:SupportIMAPIdleCommand
			' Connect and log in to IMAP 
			Dim client As New ImapClient("imap.domain.com", "username", "password")

			Dim manualResetEvent As New ManualResetEvent(False)
			Dim eventArgs As ImapMonitoringEventArgs = Nothing
			client.StartMonitoring(Sub(sender As Object, e As ImapMonitoringEventArgs) 
			eventArgs = e
			manualResetEvent.[Set]()

End Sub)
			Thread.Sleep(2000)
			Dim smtpClient As New SmtpClient("exchange.aspose.com", "username", "password")
            smtpClient.Send(New MailMessage("from@aspose.com", "to@aspose.com", "EMAILNET-34875 - " + Guid.NewGuid().ToString(), "EMAILNET-34875 Support for IMAP idle command"))
			manualResetEvent.WaitOne(10000)
			manualResetEvent.Reset()
			Console.WriteLine(eventArgs.NewMessages.Length)
			Console.WriteLine(eventArgs.DeletedMessages.Length)
			client.StopMonitoring("Inbox")
            smtpClient.Send(New MailMessage("from@aspose.com", "to@aspose.com", "EMAILNET-34875 - " + Guid.NewGuid().ToString(), "EMAILNET-34875 Support for IMAP idle command"))
			manualResetEvent.WaitOne(5000)
			' ExEnd:SupportIMAPIdleCommand
		End Sub
	End Class
End Namespace