// Demonstrates forwarding with the task-based API.
//
// ForwardAsync takes an SmtpForward parameter object: the envelope sender, one or more
// recipients, and the message - a MailMessage, or a stream with the raw message, so a
// saved .eml file can be forwarded without loading it first. The message goes out as it
// is, addressed to the new envelope recipients - unlike composing a new message that
// quotes the original.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients.Smtp.Models;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ForwardMessageAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var emlPath = Data.Smtp/"Message.eml";

            using (var smtpClient = ClientBuilder.Smtp(AuthType.Basic))
            {
                IAsyncSmtpClient client = smtpClient;
                var self = smtpClient.Username;

                // A message loaded into a MailMessage.
                await client.ForwardAsync(SmtpForward.Create()
                    .SetSender(self)
                    .AddRecipient(self)
                    .SetMessage(MailMessage.Load(emlPath)));
                Console.WriteLine($"Forwarded {Path.GetFileName(emlPath)} as a MailMessage to {self}.");

                // The same file as a raw stream.
                using (var stream = File.OpenRead(emlPath))
                {
                    await client.ForwardAsync(SmtpForward.Create()
                        .SetSender(self)
                        .AddRecipient(self)
                        .SetMessage(stream));
                }
                Console.WriteLine($"Forwarded {Path.GetFileName(emlPath)} as a stream to {self}.");
            }
        }
    }
}
