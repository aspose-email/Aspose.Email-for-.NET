using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class ConnectExchangeServerUsingIMAP
    {
        public static void Run()
        {
            // ExStart:ConnectExchangeServerUsingIMAP
            // Connect to Exchange Server using ImapClient class
            ImapClient imapClient = new ImapClient("ex07sp1", "Administrator", "Evaluation1");
            imapClient.SecurityOptions = SecurityOptions.Auto;

            // Select the Inbox folder
            imapClient.SelectFolder(ImapFolderInfo.InBox);

            // Get the list of messages
            ImapMessageInfoCollection msgCollection = imapClient.ListMessages();
            foreach (ImapMessageInfo msgInfo in msgCollection)
            {
                Console.WriteLine(msgInfo.Subject);
            }
            // Disconnect from the server
            imapClient.Dispose();
            // ExEnd:ConnectExchangeServerUsingIMAP
        }
    }
}