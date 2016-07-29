using System;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SSLEnabledIMAPServer
    {
        public static void Run()
        {
            //ExStart:SSLEnabledIMAPServer
            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();
            // Specify host, username and password, Port and SecurityOptions for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                Console.WriteLine("Logged in to the IMAP server");
                // Disconnect to the remote IMAP server
                client.Dispose();
                Console.WriteLine("Disconnected from the IMAP server");
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
            //ExEnd:SSLEnabledIMAPServer
            Console.WriteLine(Environment.NewLine + "Connected to IMAP server with SSL.");
        }
    }
}
