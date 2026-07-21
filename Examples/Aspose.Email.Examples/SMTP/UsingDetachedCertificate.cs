using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class UsingDetachedCertificate
    {
        public static void Run()
        {
            string privateCertFile = Data.Smtp + "MartinCertificate.pfx";
            X509Certificate2 privateCert = new X509Certificate2(privateCertFile, "anothertestaccount");

            MailMessage msg = new MailMessage("user@domain.com", "receiver@domain.com", "subject:Signed message only by AE", "body:Test Body of signed message by AE");

            MailMessage signed = msg.AttachSignature(privateCert, true);
            SmtpClient smtp = GetSmtpClient();
            
            try
            {
                smtp.Send(signed);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static SmtpClient GetSmtpClient()
        {
 	        SmtpClient client = new SmtpClient("smtp.domain.com", "user@domain.com", "password");
            client.Port = 25;
            client.SecurityOptions = SecurityOptions.SSLAuto;
            return client;
        }
    }
}
