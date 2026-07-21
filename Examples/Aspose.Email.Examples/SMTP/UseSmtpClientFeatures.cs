using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class UseSmtpClientFeatures
    {
        public static void Run()
        {
            // Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient("smtp.domain.com", 25);

            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            client.Timeout = 10000;
        }
    }
}
