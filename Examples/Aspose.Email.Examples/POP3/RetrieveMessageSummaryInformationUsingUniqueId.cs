using System;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.POP3
{
    class RetrieveMessageSummaryInformationUsingUniqueId
    {
        public static void Run()
        {
            string uniqueId = "unique id of a message from server";
            Pop3Client client = new Pop3Client("host.domain.com", 456, "username", "password");
            client.SecurityOptions = SecurityOptions.Auto;
            Pop3MessageInfo messageInfo = client.GetMessageInfo(uniqueId);

            if (messageInfo != null)
            {
                Console.WriteLine(messageInfo.Subject);
                Console.WriteLine(messageInfo.Date);
            }
        }
    }
}