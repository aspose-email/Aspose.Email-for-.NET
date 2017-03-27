using Aspose.Email.Clients.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ListingMIMEMessageIdInImapMessageInfo
    {
        public static void Run()
        {
            // ExStart:ListingMIMEMessageIdInImapMessageInfo
            ImapClient client = new ImapClient();
            client.Host = "domain.com";
            client.Username = "username";
            client.Password = "password";
                
            try
            {
                ImapMessageInfoCollection messageInfoCol = client.ListMessages("Inbox");
                foreach (ImapMessageInfo info in messageInfoCol)
                {
                    // Display MIME Message ID
                    Console.WriteLine("Message Id = " + info.MessageId);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:ListingMIMEMessageIdInImapMessageInfo
        }
    }
}
