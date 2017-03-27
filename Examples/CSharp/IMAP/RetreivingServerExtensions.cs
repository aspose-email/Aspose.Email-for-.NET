using System;
using Aspose.Email.Clients.Imap;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class RetreivingServerExtensions
    {
        public static void Run()
        {
            // ExStart:RetreivingServerExtensions
            // Connect and log in to IMAP
            ImapClient client = new ImapClient("imap.gmail.com", "username", "password");
            string[] getCapabilities = client.GetCapabilities();
            foreach (string getCap in getCapabilities)
            {
                Console.WriteLine(getCap);
            }
            // ExEnd:RetreivingServerExtensions
        }
    }
}