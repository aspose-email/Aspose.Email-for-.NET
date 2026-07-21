using System;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.EWS
{
    class RetrieveExtraParametersAsSummaryInformation
    {
        public static void Run()
        {
            try
            {
                using (ImapClient client = new ImapClient("host.domain.com", "username", "password"))
                {
                    MailMessage message = new MailMessage("from@domain.com", "to@doman.com", "EMAILNET-38466 - " + Guid.NewGuid().ToString(), "EMAILNET-38466 Add extra parameters for UID FETCH command");

                    // append the message to the server
                    string uid = client.AppendMessage(message);

                    // wait for the message to be appended
                    Thread.Sleep(5000);

                    // Define properties to be fetched from server along with the message
                    string[] messageExtraFields = new string[] { "X-GM-MSGID", "X-GM-THRID" };

                    // retreive the message summary information using it's UID
                    ImapMessageInfo messageInfoUID = client.ListMessage(uid, messageExtraFields);

                    // retreive the message summary information using it's sequence number
                    ImapMessageInfo messageInfoSeqNum = client.ListMessage(1, messageExtraFields);

                    // List messages in general from the server based on the defined properties
                    ImapMessageInfoCollection messageInfoCol = client.ListMessages(messageExtraFields);

                    ImapMessageInfo messageInfoFromList = messageInfoCol[0];

                    // verify that the parameters are fetched in the summary information
                    foreach (string paramName in messageExtraFields)
                    {
                        Console.WriteLine(messageInfoFromList.ExtraParameters.ContainsKey(paramName));
                        Console.WriteLine(messageInfoUID.ExtraParameters.ContainsKey(paramName));
                        Console.WriteLine(messageInfoSeqNum.ExtraParameters.ContainsKey(paramName));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}