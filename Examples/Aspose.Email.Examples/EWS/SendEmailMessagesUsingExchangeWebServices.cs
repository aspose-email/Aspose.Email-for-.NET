// Demonstrates sending a message through Exchange.
//
// IEWSClient.Send submits the message through the mailbox, so it is sent with the
// mailbox's identity and a copy lands in its Sent Items - no separate SMTP server or
// credentials are needed. The message goes to the configured mailbox itself.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class SendEmailMessagesUsingExchangeWebServices
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var message = new MailMessage
                {
                    From = self,
                    To = self,
                    Subject = "Sent through Exchange Web Services",
                    HtmlBody = "<h3>Hello</h3><p>This message was sent with IEWSClient.Send.</p>"
                };

                client.Send(message);
                Console.WriteLine($"Sent '{message.Subject}' to {self}.");
            }
        }
    }
}
