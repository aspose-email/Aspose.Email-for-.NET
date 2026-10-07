// Demonstrates how to connect to an IMAP server.
//
// Creating an ImapClient stores the server address and the credentials; the client
// connects and signs in when the first command needs the server, here SelectFolder.
// Disposing the client logs out and closes the connection, so keep it in a using block.
// The host, port and account come from clientsettings.json through ClientBuilder.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ConnectingWithIMAPServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                Console.WriteLine($"Server:   {client.Host}:{client.Port}");
                Console.WriteLine($"Account:  {client.Username}");
                Console.WriteLine($"Security: {client.SecurityOptions}");

                client.SelectFolder(ImapFolderInfo.InBox);

                Console.WriteLine($"\nConnected, connection is {client.ConnectionState}.");
                Console.WriteLine($"The Inbox holds {client.CurrentFolder.TotalMessageCount} message(s).");
            }

            Console.WriteLine("Disconnected.");
        }
    }
}
