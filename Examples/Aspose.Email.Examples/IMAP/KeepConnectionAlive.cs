// Demonstrates how to keep a long-lived IMAP session from being dropped as idle.
//
// Servers disconnect clients that stay silent too long (RFC 3501 allows 30 minutes, and
// firewalls or NAT routers often cut in sooner). Noop sends a harmless NOOP command that
// resets the idle timer. ConnectionCheckupPeriod sets how often the client itself makes
// sure its connection is still usable.

using System;
using System.Threading;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class KeepConnectionAlive
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.ConnectionCheckupPeriod = 60000;  // milliseconds; the default is 5 minutes

                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"{DateTime.Now:T}  Inbox selected, connection is {client.ConnectionState}");

                for (var i = 0; i < 3; i++)
                {
                    // Stand-in for local work that keeps the connection quiet for a while.
                    Thread.Sleep(TimeSpan.FromSeconds(10));

                    client.Noop();
                    Console.WriteLine($"{DateTime.Now:T}  NOOP sent, connection is {client.ConnectionState}");
                }
            }
        }
    }
}
