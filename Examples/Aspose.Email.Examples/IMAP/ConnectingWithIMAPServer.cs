using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    class ConnectingWithIMAPServer
    {
        public static void Run()
        {
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient("localhost", "user", "password");

            try
            {
                // Disconnect to the remote IMAP server
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            Console.WriteLine(Environment.NewLine + "Connected to IMAP server.");            
        }
    }
}
