using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class RetrieveEmailViaPop3ClientProxyServer
    {
        public static void Run()
        {
            // ExStart:RetrieveEmailViaPop3ClientProxyServer
            Pop3Client client = new Pop3Client("pop.domain.com", "username", "password");
          
            // Set proxy address, Port and Proxy
            string proxyAddress = "192.168.203.142";
            int proxyPort = 1080;
            SocksProxy proxy = new SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5);
            client.SocksProxy = proxy;

            try
            {
                Pop3MailboxInfo mailboxInfo = client.GetMailboxInfo();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
            // ExEnd:RetrieveEmailViaPop3ClientProxyServer
        }
    }
}