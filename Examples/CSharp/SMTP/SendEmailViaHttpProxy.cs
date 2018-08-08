using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.SMTP
{
    public class SendEmailViaHttpProxy
    {
        public static void Run()
        {
            //ExStart: SendEmailViaHttpProxy
            HttpProxy proxy = new HttpProxy("18.222.124.59", 8080);
            using (SmtpClient client = new SmtpClient("host", 587, "username", "password"))
            {
                client.Proxy = proxy;
                client.Send(new MailMessage(
                    "from@domain.com",
                    "to@domain.com",
                    "NETWORKNET-34226 - " + Guid.NewGuid().ToString(),
                    "NETWORKNET-34226 Implement socks proxy protocol for versions 4, 4a, 5 (only Username/Password authentication)"));
            }
            //ExEnd: SendEmailViaHttpProxy
        }
    }
}
