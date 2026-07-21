// Demonstrates how to load an EML file and print all its email headers with their values.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ExtractingEmailHeaders
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"email-headers.eml", new EmlLoadOptions());

            var index = 0;
            foreach (var header in eml.Headers)
                Console.WriteLine($"{header}: {eml.Headers.Get(index++)}");
        }
    }
}
