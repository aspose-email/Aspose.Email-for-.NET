using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Clients.Pop3;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class Pop3FetchGroupMessages
    {
        public static void Run()
        {
            //ExStart:1
            // Create an instance of the Pop3Client class
            Pop3Client pop3Client = new Pop3Client();
            pop3Client.Host = "<HOST>";
            pop3Client.Port = 995;
            pop3Client.Username = "<USERNAME>";
            pop3Client.Password = "<PASSWORD>";

            Pop3MessageInfoCollection messageInfoCol = pop3Client.ListMessages();
            Console.WriteLine("ListMessages Count: " + messageInfoCol.Count);
            int[] sequenceNumberAr = messageInfoCol.Select((Pop3MessageInfo mi) => mi.SequenceNumber).ToArray();
            string[] uniqueIdAr = messageInfoCol.Select((Pop3MessageInfo mi) => mi.UniqueId).ToArray();

            IList<MailMessage> fetchedMessagesBySNumMC = pop3Client.FetchMessages(sequenceNumberAr);
            Console.WriteLine("FetchMessages-sequenceNumberAr Count: " + fetchedMessagesBySNumMC.Count);

            IList<MailMessage> fetchedMessagesByUidMC = pop3Client.FetchMessages(uniqueIdAr);
            Console.WriteLine("FetchMessages-uniqueIdAr Count: " + fetchedMessagesByUidMC.Count);
            //ExEnd: 1

            Console.WriteLine("Pop3FetchGroupMessages executed successfully.");
        }
    }
}
