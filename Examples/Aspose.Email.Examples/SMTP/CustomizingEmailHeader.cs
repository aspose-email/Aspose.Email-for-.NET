// Demonstrates how to send a message with a custom header.
//
// Custom headers - by convention starting with "X-" - travel with the message to the
// recipient, where filters and programs can read them, for example to recognise mail
// sent by your application. XMailer names the sending program. CustomizingEmailHeaders
// builds a similar message and saves it instead.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class CustomizingEmailHeader
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                var message = new MailMessage(client.Username, client.Username,
                    "Message with custom headers", "Look at the headers of this message.")
                {
                    XMailer = "Aspose.Email Examples"
                };

                message.ReplyToList.Add(client.Username);
                message.Headers.Add("X-Campaign-Id", "spring-2026");
                message.Headers.Add("X-Secret-Header", "mystery");

                client.Send(message);
                Console.WriteLine($"Sent to {client.Username} with X-Campaign-Id and X-Secret-Header.");
            }
        }
    }
}
