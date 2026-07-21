// Demonstrates how to check whether an email message is a bounce notification
// and inspect the bounce details such as action, recipient, reason, and status.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class CheckBouncedMessage
    {
        public static void Run()
        {
            var result = MailMessage.Load(Data.Email/"failed1.msg").CheckBounced();

            Console.WriteLine($"IsBounced: {result.IsBounced}");
            Console.WriteLine($"Action: {result.Action}");
            Console.WriteLine($"Recipient: {result.Recipient}");
            Console.WriteLine($"Reason: {result.Reason}");
            Console.WriteLine($"Status: {result.Status}");
            Console.WriteLine($"OriginalMessage ToAddress 1: {result.OriginalMessage.To[0].Address}");
        }
    }
}
