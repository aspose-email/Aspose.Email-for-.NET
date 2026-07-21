using System;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.IMAP
{
    class SupportIMAPIdleCommand
    {
        public static void Run()
        {
            // Connect and log in to IMAP 
            ImapClient client = new ImapClient("imap.domain.com", "username", "password");

            ManualResetEvent manualResetEvent = new ManualResetEvent(false);
            ImapMonitoringEventArgs eventArgs = null;
            client.StartMonitoring(delegate(object sender, ImapMonitoringEventArgs e)
            {
                eventArgs = e;
                manualResetEvent.Set();
            }, null);
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
        }
    }
}