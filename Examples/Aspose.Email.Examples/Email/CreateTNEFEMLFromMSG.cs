// Demonstrates how to convert an MSG file to a TNEF-encoded EML message
// using MailConversionOptions.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class CreateTnefEmlFromMsg
    {
        public static void Run()
        {
            var mapiMsg = MapiMessage.Load(Data.Email/"Message.msg");

            var mailConversionOptions = new MailConversionOptions { ConvertAsTnef = true };
            var eml = mapiMsg.ToMailMessage(mailConversionOptions);

            Console.WriteLine($"Message is TNEF-encoded: {eml.OriginalIsTnef}");
        }
    }
}
