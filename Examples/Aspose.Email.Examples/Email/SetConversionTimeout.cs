// Demonstrates how to bound the time a MapiMessage-to-MailMessage conversion may
// take, and how to notice that the limit was hit.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class SetConversionTimeout
    {
        public static void Run()
        {
            var mapiMessage = MapiMessage.Load(Data.Email/"Message.msg");

            var isTimedOut = false;

            var options = new MailConversionOptions { Timeout = 5000 };
            options.TimeoutReached += (sender, args) =>
            {
                var subject = (sender as MailMessage)?.Subject;
                Console.WriteLine($"Conversion timed out for '{subject}'.");
                isTimedOut = true;
            };

            var mailMessage = mapiMessage.ToMailMessage(options);

            Console.WriteLine($"Timeout:   {options.Timeout} ms");
            Console.WriteLine($"Converted: {mailMessage.Subject}");
            Console.WriteLine($"Timed out: {isTimedOut}");
        }
    }
}
