using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SSLEnabledIMAPServer
    {
        public static void Run()
        {
            //ExStart:SSLEnabledIMAPServer
            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient("imap.domain.com", 993, "user@domain.com", "pwd");
            
            // Set the security mode to implicit
            client.SecurityOptions = SecurityOptions.SSLImplicit;
            //ExEnd:SSLEnabledIMAPServer

            try
            {
                Console.WriteLine("Logged in to the IMAP server");
                // Disconnect to the remote IMAP server
                client.Dispose();
                Console.WriteLine("Disconnected from the IMAP server");
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            Console.WriteLine(Environment.NewLine + "Connected to IMAP server with SSL.");
        }
    }
}
