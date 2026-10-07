// Demonstrates how to find out what went wrong when the server refuses a message.
//
// - A refused recipient raises SmtpFailedRecipientException, or
//   SmtpFailedRecipientsException with one inner exception per refused address. Each
//   carries the address (FailedRecipient) and the server's reply code (StatusCode).
// - When a batch is sent, SmtpException.OperationDetails tells which messages were
//   sent, which failed and why, and which were not attempted.
//
// The example sends to a made-up address in your own domain. Many servers refuse it at
// once; others accept it and send a bounce message to you later, in which case nothing
// is reported here.

using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class HandleSmtpSendErrors
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
                var self = client.Username;
                var unknown = "no-such-user-" + Guid.NewGuid().ToString("N").Substring(0, 8) + "@" + new MailAddress(self).Host;

                Console.WriteLine($"1. One message to {unknown}:");
                try
                {
                    client.Send(new MailMessage(self, unknown, "Undeliverable", "Body"));
                    Console.WriteLine("   accepted by the server (a bounce may follow).");
                }
                catch (SmtpFailedRecipientsException ex)
                {
                    foreach (var inner in ex.InnerExceptions)
                        Console.WriteLine($"   refused {inner.FailedRecipient}: {inner.StatusCode}");
                }
                catch (SmtpFailedRecipientException ex)
                {
                    Console.WriteLine($"   refused {ex.FailedRecipient}: {ex.StatusCode}");
                }
                catch (SmtpException ex)
                {
                    Console.WriteLine($"   failed with {ex.StatusCode}: {ex.Message}");
                }

                Console.WriteLine("\n2. A batch of three, the second one to the unknown address:");
                var batch = new List<MailMessage>
                {
                    new MailMessage(self, self, "Batch message 1", "Body"),
                    new MailMessage(self, unknown, "Batch message 2", "Body"),
                    new MailMessage(self, self, "Batch message 3", "Body")
                };

                try
                {
                    client.Send(batch);
                    Console.WriteLine("   all accepted by the server.");
                }
                catch (SmtpException ex)
                {
                    Console.WriteLine($"   {ex.Message}");

                    var details = ex.OperationDetails;
                    if (details == null)
                        return;

                    foreach (var message in details.Succeeded)
                        Console.WriteLine($"   sent:        {message.Subject}");
                    foreach (var failure in details.Failed)
                        Console.WriteLine($"   failed:      {failure.Key.Subject} - {failure.Value.Message}");
                    foreach (var message in details.NotHandled)
                        Console.WriteLine($"   not handled: {message.Subject}");
                }
            }
        }
    }
}
