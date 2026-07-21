// Demonstrates how to detect whether a loaded EML message was originally TNEF-encoded.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class DetectMessageIsTnef
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"tnefEml1.eml");
            Console.WriteLine($"Is input EML originally TNEF? {eml.OriginalIsTnef}");
        }
    }
}
