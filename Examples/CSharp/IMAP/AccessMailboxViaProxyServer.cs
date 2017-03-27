using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using System;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class AccessMailboxViaProxyServer
    {
        public static void Run()
        {
            // ExStart:AccessMailboxViaProxyServer
            // Connect and log in to IMAP and set SecurityOptions
            ImapClient client = new ImapClient("imap.domain.com", "username", "password");
            client.SecurityOptions = SecurityOptions.Auto;
            
            string proxyAddress = "192.168.203.142"; // proxy address
            int proxyPort = 1080; // proxy port
            SocksProxy proxy = new SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5);

            // Set the proxy
            client.SocksProxy = proxy;
           
            try
            {
                client.SelectFolder("Inbox");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:AccessMailboxViaProxyServer
        }
    }
}