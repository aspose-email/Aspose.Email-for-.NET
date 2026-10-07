// Demonstrates how to send from several threads with one SmtpClient.
//
// An SMTP connection carries one message at a time, so threads sharing it would wait
// for each other. CreateConnection opens an extra, independent connection; pass it to
// the Send overloads that take an IConnection and the message goes over that
// connection. Dispose it when done. SendWithMultiConnection shows the automatic variant,
// where the client spreads one big batch over several connections by itself.

using System;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendOverDedicatedConnections
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

                // Three workers, each sending two messages over a connection of its own.
                var workers = Enumerable.Range(1, 3).Select(worker => Task.Run(() =>
                {
                    using (var connection = client.CreateConnection())
                    {
                        for (var i = 1; i <= 2; i++)
                        {
                            client.Send(connection, new MailMessage(self, self,
                                $"Worker {worker}, message {i}", "Body"));
                        }

                        return $"Worker {worker} sent 2 message(s) over connection #{connection.ConnectionId}.";
                    }
                })).ToArray();

                Task.WaitAll(workers);

                foreach (var worker in workers)
                    Console.WriteLine(worker.Result);
            }
        }
    }
}
