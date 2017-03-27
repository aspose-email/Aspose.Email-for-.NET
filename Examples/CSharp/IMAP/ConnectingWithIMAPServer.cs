using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ConnectingWithIMAPServer
    {
        public static void Run()
        {
            // ExStart:ConnectingWithIMAPServer
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient("localhost", "user", "password");
            // ExEnd:ConnectingWithIMAPServer

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
