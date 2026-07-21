// Demonstrates how to load an EML file and display its basic header properties.

using System;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExposeProperties
    {
        public static void Run()
        {
            var msg = MailMessage.Load(Data.Mapi/"Message.eml", new EmlLoadOptions());

            Console.WriteLine("Subject: " + (msg.Subject ?? "no subject"));
            Console.WriteLine("From: " + (msg.From != null ? msg.From.ToString() : "No sender"));
            Console.WriteLine("To: " + (msg.To != null ? msg.To.ToString() : "No recipients"));
            Console.WriteLine("CC: " + (msg.CC != null ? msg.CC.ToString() : "No CC recipients"));
        }
    }
}
