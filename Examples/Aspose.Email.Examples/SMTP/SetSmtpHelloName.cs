// Demonstrates how to choose the name the client announces in its EHLO/HELO greeting.
//
// The server records this name in the Received header and some servers check it: a
// greeting that does not match the sending machine's DNS name can count against the
// message in spam filtering. Set HelloMessage to the fully qualified name of the host
// the program runs on when the default is not right.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SetSmtpHelloName
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
                client.HelloMessage = "mailer.example.com";

                client.Send(new MailMessage(client.Username, client.Username,
                    "EHLO name test", $"Sent with the EHLO name '{client.HelloMessage}'."));

                Console.WriteLine($"Sent to {client.Username}, greeting the server as '{client.HelloMessage}'.");
                Console.WriteLine("Look for that name in the Received header of the delivered message.");
            }
        }
    }
}
