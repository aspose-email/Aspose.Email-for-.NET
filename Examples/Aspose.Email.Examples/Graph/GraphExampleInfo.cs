// Shared helper for the Graph examples. Not an example itself - it has no Run(), so
// the example runner does not list it.

using System;

namespace Aspose.Email.Examples.Graph
{
    internal static class GraphExampleInfo
    {
        // Every Graph example calls this when clientsettings.json has no application
        // registration, so that running one offline explains what is missing instead of
        // failing with an authentication error.
        internal static void PrintNotConfigured()
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Microsoft Graph is not configured, so this example did nothing.");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Fill in the \"Graph\" section of clientsettings.json:");
            Console.WriteLine("  TenantId     - the directory (tenant) id of your Entra ID tenant");
            Console.WriteLine("  ClientId     - the application (client) id of your app registration");
            Console.WriteLine("  ClientSecret - a client secret of that registration");
            Console.WriteLine("  MailboxId    - the user principal name of the mailbox to work on");
            Console.WriteLine();
            Console.WriteLine("The registration needs application permissions for what it uses, for");
            Console.WriteLine("example Mail.ReadWrite, Mail.Send, Contacts.ReadWrite, Calendars.ReadWrite");
            Console.WriteLine("or Tasks.ReadWrite, each granted admin consent.");
        }
    }
}
