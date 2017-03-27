using System;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class RetrieveExtraParametersAsSummaryInformation
    {
        public static void Run()
        {
            try
            {
                // ExStart:RetrieveExtraParametersAsSummaryInformation
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
                // ExEnd:RetrieveExtraParametersAsSummaryInformation
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}