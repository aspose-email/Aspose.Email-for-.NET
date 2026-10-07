// Shared helper for the SMTP examples. Not an example itself - it has no Run(), so the
// example runner does not list it.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SmtpExampleInfo
    {
        // The SMTP examples that read clientsettings.json call this when no server is
        // configured, so that running one offline explains what is missing instead of
        // waiting for a connection that cannot succeed.
        internal static void PrintNotConfigured()
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("SMTP is not configured, so this example did nothing.");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Fill in the \"Smtp\" section of clientsettings.json:");
            Console.WriteLine("  HostName - the SMTP server, for example smtp.example.com");
            Console.WriteLine("  Port     - 587 for STARTTLS, 465 for implicit TLS");
            Console.WriteLine("  UserName - the account to sign in with, as an e-mail address;");
            Console.WriteLine("             the examples send their test messages to it");
            Console.WriteLine("  Password - its password, or an app password where the provider");
            Console.WriteLine("             requires one");
        }
    }
}
