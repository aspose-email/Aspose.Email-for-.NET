using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class RetreiveServerExtensions
    {
        public static void Run()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", "user@gmail.com", "password");
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            client.Port = 587;

            try
            {
                string[] caps = client.GetCapabilities();

                foreach (string str in caps)
                    Console.WriteLine(str);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
