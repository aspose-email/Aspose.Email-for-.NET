using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class UsingDetachedCertificate
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_SMTP();

            // ExStart:UsingDetachedCertificate
            string privateCertFile = dataDir + "MartinCertificate.pfx";
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
            // ExEnd:UsingDetachedCertificate
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
