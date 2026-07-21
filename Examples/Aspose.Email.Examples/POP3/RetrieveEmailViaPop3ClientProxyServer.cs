using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.POP3
{
    class RetrieveEmailViaPop3ClientProxyServer
    {
        public static void Run()
        {
            Pop3Client client = new Pop3Client("pop.domain.com", "username", "password");
          
            // Set proxy address, Port and Proxy
            string proxyAddress = "192.168.203.142";
            int proxyPort = 1080;
            SocksProxy proxy = new SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5);
            client.Proxy = proxy;

            try
            {
                Pop3MailboxInfo mailboxInfo = client.GetMailboxInfo();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}