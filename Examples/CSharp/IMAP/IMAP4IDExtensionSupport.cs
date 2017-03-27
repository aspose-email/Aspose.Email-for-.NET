using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class IMAP4IDExtensionSupport
    {
        public static void Run()
        {
            //ExStart:IMAP4IDExtensionSupport
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
            //ExEnd:IMAP4IDExtensionSupport
        }
    }
}
