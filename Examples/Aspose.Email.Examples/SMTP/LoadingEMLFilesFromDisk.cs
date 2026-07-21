using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    class LoadingEMLFilesFromDisk
    {
        public static void Run()
        {
            // Load an EML file in MailMessage class
            MailMessage message = MailMessage.Load(Data.Smtp + "test.eml");

            // Send this message using SmtpClient
            SmtpClient client = new SmtpClient("host", "username", "password");
            
            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
        }
    }
}
