using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class SSLEnabledPOP3Server
    {
        public static void Run()
        {
            //ExStart:SSLEnabledPOP3Server
            // Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();
            // Specify host, username and password, Port and  SecurityOptions for your client
            client.Host = "pop.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.Auto; 
            Console.WriteLine(Environment.NewLine + "Connecting to POP3 server using SSL.");
            //ExEnd:SSLEnabledPOP3Server
        }
    }
}
