// Demonstrates how to load an email and check whether it is a bounced message,
// displaying the bounce action, recipient, and status.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ProcessBouncedMsgs
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"test.eml");
            var result = eml.CheckBounced();

            Console.WriteLine($"IsBounced: {result.IsBounced}");
            Console.WriteLine($"Action: {result.Action}");
            Console.WriteLine($"Recipient: {result.Recipient}");
        }
    }
}
