using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.POP3
{
    class GetServerExtensionsUsingPop3Client
    {
        public static void Run()
        {
            // Connect and log in to POP3
            Pop3Client client = new Pop3Client("pop.gmail.com", "username", "password");
            client.SecurityOptions = SecurityOptions.Auto;
            client.Port = 993;
            string[] getCaps = client.GetCapabilities();
            foreach (string item in getCaps)
            {
                Console.WriteLine(item);
            }
        }
    }
}