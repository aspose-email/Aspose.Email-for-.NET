using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.IMAP
{
    class IMAP4IDExtensionSupport
    {
        public static void Run()
        {
            using (ImapClient client = new ImapClient("imap.gmail.com", 993, "username", "password"))
            {
                // Set SecurityOptions
                client.SecurityOptions = SecurityOptions.Auto;
                Console.WriteLine(client.IdSupported.ToString());

                ImapIdentificationInfo serverIdentificationInfo1 = client.IntroduceClient();
                ImapIdentificationInfo serverIdentificationInfo2 = client.IntroduceClient(ImapIdentificationInfo.DefaultValue);

                // Display ImapIdentificationInfo properties
                Console.WriteLine(serverIdentificationInfo1.ToString(), serverIdentificationInfo2);
                Console.WriteLine(serverIdentificationInfo1.Name);
                Console.WriteLine(serverIdentificationInfo1.Vendor);
                Console.WriteLine(serverIdentificationInfo1.SupportUrl);
                Console.WriteLine(serverIdentificationInfo1.Version);
            }
        }
    }
}
