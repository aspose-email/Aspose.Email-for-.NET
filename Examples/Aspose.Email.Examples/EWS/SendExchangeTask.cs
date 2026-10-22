// Demonstrates sending a task created in code.
//
// An ExchangeTask can be saved as an Outlook .msg; loading that with
// MsgLoadOptions.PreserveTnefAttachments gives a MailMessage that carries the task, ready
// to send. The task goes to the configured mailbox itself, where Outlook shows it as a
// task.

using System;
using System.IO;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class SendExchangeTask
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;

            var task = new ExchangeTask
            {
                Subject = "Review the contract",
                Status = ExchangeTaskStatus.NotStarted,
                StartDate = DateTime.Today,
                DueDate = DateTime.Today.AddDays(2),
                Body = "Please review the attached contract by Friday."
            };

            using (var stream = new MemoryStream())
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                task.Save(stream, new TaskMsgSaveOptions());
                stream.Position = 0;

                var message = MailMessage.Load(stream, new MsgLoadOptions { PreserveTnefAttachments = true });
                message.From = self;
                message.To.Clear();
                message.To.Add(self);

                client.Send(message);
                Console.WriteLine($"Sent the task '{task.Subject}' to {self}.");
            }
        }
    }
}
