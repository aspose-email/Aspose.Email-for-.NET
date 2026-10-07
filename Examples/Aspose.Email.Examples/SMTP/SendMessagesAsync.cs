// Demonstrates sending with the task-based API.
//
// IAsyncSmtpClient.SendAsync takes an SmtpSend parameter object from the
// Aspose.Email.Clients.Smtp.Models namespace: SmtpSend.Create() starts it, AddMessage /
// AddMessages collect what to send - a ready MailMessage or just from, to, subject and
// body - and SetCancellationToken lets the caller abandon a send that takes too long.
// SmtpClient implements IAsyncSmtpClient, so any client can be used this way.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients.Smtp.Models;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendMessagesAsync
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

            using (var smtpClient = ClientBuilder.Smtp(AuthType.Basic))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(1)))
            {
                IAsyncSmtpClient client = smtpClient;
                var self = smtpClient.Username;

                try
                {
                    var valid = await client.ValidateCredentialsAsync(token: cancellation.Token);
                    Console.WriteLine($"Credentials accepted: {valid}");

                    var reports = Enumerable.Range(1, 3)
                        .Select(i => new MailMessage(self, self, $"Regional report {i}", "Report attached."));

                    var parameters = SmtpSend.Create()
                        .AddMessage(self, self, "Daily summary", "Three reports follow.")
                        .AddMessages(reports)
                        .SetCancellationToken(cancellation.Token);

                    await client.SendAsync(parameters);
                    Console.WriteLine($"Sent 4 message(s) to {self}.");
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine("Gave up: sending took longer than a minute.");
                }
            }
        }
    }
}
