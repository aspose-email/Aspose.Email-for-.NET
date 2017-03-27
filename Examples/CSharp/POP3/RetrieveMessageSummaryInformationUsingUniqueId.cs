using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class RetrieveMessageSummaryInformationUsingUniqueId
    {
        public static void Run()
        {
            // ExStart:RetrieveMessageSummaryInformationUsingUniqueId
            string uniqueId = "unique id of a message from server";
            Pop3Client client = new Pop3Client("host.domain.com", 456, "username", "password");
            client.SecurityOptions = SecurityOptions.Auto;
            Pop3MessageInfo messageInfo = client.GetMessageInfo(uniqueId);

            if (messageInfo != null)
            {
                Console.WriteLine(messageInfo.Subject);
                Console.WriteLine(messageInfo.Date);
            }
            // ExEnd:RetrieveMessageSummaryInformationUsingUniqueId
        }
    }
}