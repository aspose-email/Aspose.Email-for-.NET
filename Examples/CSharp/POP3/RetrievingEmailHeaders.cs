using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class RetrievingEmailHeaders
    {
        public static void Run()
        {
            // ExStart:RetrievingEmailHeaders
            // Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            // Specify host, username. password, Port and SecurityOptions for your client
            client.Host = "pop.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                HeaderCollection headers = client.GetMessageHeaders(1);
                for (int i = 0; i < headers.Count; i++)
                {
                    // Display key and value in the header collection
                    Console.Write(headers.Keys[i]);
                    Console.Write(" : ");
                    Console.WriteLine(headers.Get(i));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                client.Dispose();
            }
            // ExEnd:RetrievingEmailHeaders
            Console.WriteLine(Environment.NewLine + "Displayed header information from emails using POP3 ");
        }
    }
}
