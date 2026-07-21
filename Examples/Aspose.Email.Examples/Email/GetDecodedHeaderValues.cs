// Demonstrates how to retrieve and decode a specific email header value
// from an EML message using GetDecodedValue.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class GetDecodedHeaderValues
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"emlWithHeaders.eml");
            Console.WriteLine(eml.Headers.GetDecodedValue("Thread-Topic"));
        }
    }
}
