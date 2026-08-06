// Demonstrates how to bound the time an MHTML conversion may take. A message whose
// body references unreachable resources can otherwise keep the save call waiting.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class SetTimeoutForMhtSaving
    {
        public static void Run()
        {
            var mailMessage = MailMessage.Load(Data.Email/"HtmlWithUrlSample.eml");

            var timedOut = false;

            var options = SaveOptions.DefaultMhtml;
            options.Timeout = 4000;
            options.TimeoutReached += (sender, e) => { timedOut = true; };

            using (var ms = new MemoryStream())
            {
                mailMessage.Save(ms, options);

                Console.WriteLine($"Timeout:  {options.Timeout} ms");
                Console.WriteLine($"Produced: {ms.Length} byte(s)");
                Console.WriteLine($"Timed out: {timedOut}");
            }
        }
    }
}
