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

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class AddingMessageToGmailFolderUsingIMAP
    {
        public static void Run()
        {
            try
            {
                // Connect to the Gmail server
                ImapClient imap = new ImapClient("imap.gmail.com", 993, "user@gmail.com", "pwd");

                // ExStart:AddingMessageToGmailFolderUsingIMAP
                // Create a message and Subscribe to the Inbox folder ans Append the newly created message
                MailMessage message = new MailMessage("user@domain1.com", "user@domain2.com", "subject", "message");
                imap.SelectFolder(ImapFolderInfo.InBox);
                imap.SubscribeFolder(imap.CurrentFolder.Name);
                imap.AppendMessage(imap.CurrentFolder.Name, message);
                // ExEnd:AddingMessageToGmailFolderUsingIMAP
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
