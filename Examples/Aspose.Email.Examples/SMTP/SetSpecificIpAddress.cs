using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    class SetSpecificIpAddress
    {
        public static void Run()
        {
            using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.Password", SecurityOptions.Auto))
            {
                // set callback
                client.BindIPEndPoint += BindIPEndPointCallback;
                client.Noop();
            }
        }

        // get local endpoint callbak
        private static IPEndPoint BindIPEndPointCallback(IPEndPoint remoteEndPoint)
        {
            return new IPEndPoint(IPAddress.Any, 0);
        }
    }
}
