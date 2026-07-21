using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.POP3
{
    class GetEmailCountIntheMailbox
    {
        public static void Run()
        {
            Pop3Client client = new Pop3Client("pop3.server.com", "username", "password");
            try
            {
                client.SecurityOptions = SecurityOptions.Auto;
                int i = client.GetMessageCount();
                Console.WriteLine("Message count: " + i);
            }
            catch (Pop3Exception ex)
            {
                Console.WriteLine("Error:" + ex.ToString());
            }
        }
    }
}
