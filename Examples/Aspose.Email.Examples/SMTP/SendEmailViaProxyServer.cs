using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    class SendEmailViaProxyServer
    {
        public static void Run()
        {
            SmtpClient client = new SmtpClient("smtp.domain.com", "username", "password");
            client.SecurityOptions = SecurityOptions.SSLImplicit;
            string proxyAddress = "192.168.203.142"; // proxy address
            int proxyPort = 1080; // proxy port
            SocksProxy proxy = new SocksProxy(proxyAddress, proxyPort, SocksVersion.SocksV5);
            client.Proxy = proxy;
            client.Send(new MailMessage("sender@domain.com", "receiver@domain.com", "Sending Email via proxy", "Implement socks proxy protocol for versions 4, 4a, 5 (only Username/Password authentication)"));
        }
    }
}
