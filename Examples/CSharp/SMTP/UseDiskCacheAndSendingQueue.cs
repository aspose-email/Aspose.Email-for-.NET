using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class UseDiskCacheAndSendingQueue
    {
        public static void Run()
        {
            //ExStart: 1
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "<HOST>";
            smtpClient.Username = "<USERNAME>";
            smtpClient.Password = "<PASSWORD>";
            smtpClient.Port = 587;
            smtpClient.SupportedEncryption = EncryptionProtocols.Tls;
            smtpClient.SecurityOptions = SecurityOptions.SSLExplicit;

            int messageNumber = 30;
            List<MailMessage> messages = new List<MailMessage>();
            for (int i = 0; i < messageNumber; i++)
            {
                MailMessage message = new MailMessage(
                    "mactest18.email@gmail.com",
                    "aspose.test18@gmail.com",
                    "Test Message - " + Guid.NewGuid().ToString(),
                    "Use disk cache and sending queue in group SMTP send operation");
                messages.Add(message);
            }

            smtpClient.SmtpQueueLocation = @"D:\E\AsposeTestDir\queue";
            int counter = 0;
            smtpClient.SucceededQueueSending += delegate (object sender, MailMessageEventArgs arguments)
            {
                counter++;
            };
            smtpClient.FailedQueueSending += delegate (object sender, MailMessageEventArgs arguments)
            {
                counter++;
            };
            smtpClient.SendToQueue(messages);
            IAsyncResult asyncResult = smtpClient.BeginSendQueue();
            while (counter != messageNumber)
            {
                Thread.Sleep(50);
            }
            smtpClient.CancelAsyncOperation(asyncResult);
            //ExEnd: 1

            Console.WriteLine("UseDiskCacheAndSendingQueue executed successfully.");
        }
    }
}
