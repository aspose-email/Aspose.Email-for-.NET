using System;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SupportIMAPIdleCommand
    {
        public static void Run()
        {
            // ExStart:SupportIMAPIdleCommand
            // Connect and log in to IMAP 
            ImapClient client = new ImapClient("imap.domain.com", "username", "password");

            ManualResetEvent manualResetEvent = new ManualResetEvent(false);
            ImapMonitoringEventArgs eventArgs = null;
            client.StartMonitoring(delegate(object sender, ImapMonitoringEventArgs e)
            {
                eventArgs = e;
                manualResetEvent.Set();
            });
            Thread.Sleep(2000);
            SmtpClient smtpClient = new SmtpClient("exchange.aspose.com", "username", "password");
            smtpClient.Send(new MailMessage("from@aspose.com", "to@aspose.com", "EMAILNET-34875 - " + Guid.NewGuid(), "EMAILNET-34875 Support for IMAP idle command"));
            manualResetEvent.WaitOne(10000);
            manualResetEvent.Reset();
            Console.WriteLine(eventArgs.NewMessages.Length);
            Console.WriteLine(eventArgs.DeletedMessages.Length);
            client.StopMonitoring("Inbox");
            smtpClient.Send(new MailMessage("from@aspose.com", "to@aspose.com", "EMAILNET-34875 - " + Guid.NewGuid(), "EMAILNET-34875 Support for IMAP idle command"));
            manualResetEvent.WaitOne(5000);
            // ExEnd:SupportIMAPIdleCommand
        }
    }
}