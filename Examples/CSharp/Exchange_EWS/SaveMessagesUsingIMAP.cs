using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
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

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class SaveMessagesUsingIMAP
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Exchange();

            try
            {
                ImapClient imapClient = new ImapClient("ex07sp1", "Administrator", "Evaluation1");
                imapClient.SecurityOptions = SecurityOptions.Auto;

                // ExStart:SaveMessagesUsingIMAP
                // Select the Inbox folder
                imapClient.SelectFolder(ImapFolderInfo.InBox);
                // Get the list of messages
                ImapMessageInfoCollection msgCollection = imapClient.ListMessages();
                foreach (ImapMessageInfo msgInfo in msgCollection)
                {
                    // Fetch the message from inbox using its SequenceNumber from msgInfo
                    MailMessage message = imapClient.FetchMessage(msgInfo.SequenceNumber);

                    // Save the message to disc now
                    message.Save(dataDir + msgInfo.SequenceNumber + "_out.msg", SaveOptions.DefaultMsgUnicode);
                }
                // ExEnd:SaveMessagesUsingIMAP
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
