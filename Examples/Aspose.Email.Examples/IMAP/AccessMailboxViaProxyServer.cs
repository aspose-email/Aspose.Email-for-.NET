using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;

namespace Aspose.Email.Examples.IMAP
{
    class AccessMailboxViaProxyServer
    {
        public static void Run()
        {
            // Connect and log in to IMAP and set SecurityOptions
            ImapClient client = new ImapClient("imap.domain.com", "username", "password");
            client.SecurityOptions = SecurityOptions.Auto;
            
            string proxyAddress = "192.168.203.142"; // proxy address
            int proxyPort = 1080; // proxy port
            SocksProxy proxy = new SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5);

            // Set the proxy
            client.Proxy = proxy;
           
            try
            {
                client.SelectFolder("Inbox");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}