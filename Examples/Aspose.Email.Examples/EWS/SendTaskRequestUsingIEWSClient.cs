// Demonstrates sending a task request saved as an Outlook .msg file.
//
// A task request is an Outlook task packed into a message (TNEF). Loading the .msg with
// MsgLoadOptions.PreserveTnefAttachments keeps that packing intact, so the recipient's
// Outlook shows it as a task to accept or decline. The request goes to the configured
// mailbox itself.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class SendTaskRequestUsingIEWSClient
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var loadOptions = new MsgLoadOptions { PreserveTnefAttachments = true };
                var message = MailMessage.Load(Data.Ews/"task.msg", loadOptions);

                message.From = self;
                message.To.Clear();
                message.To.Add(self);

                client.Send(message);
                Console.WriteLine($"Sent the task request '{message.Subject}' to {self}.");
            }
        }
    }
}
