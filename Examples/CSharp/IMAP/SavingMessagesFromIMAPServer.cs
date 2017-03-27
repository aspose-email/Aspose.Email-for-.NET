using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SavingMessagesFromIMAPServer
    {
        public static void Run()
        {
            // ExStart:SavingMessagesFromIMAPServer
            // The path to the file directory.
            string dataDir = RunExamples.GetDataDir_IMAP();

            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient("localhost", "user", "password");

            // Select the inbox folder and Get the message info collection
            client.SelectFolder(ImapFolderInfo.InBox);
            ImapMessageInfoCollection list = client.ListMessages();

            // Download each message
            for (int i = 0; i < list.Count; i++)
            {
                // Save the message in MSG format
                MailMessage message = client.FetchMessage(list[i].UniqueId);
                message.Save(dataDir + list[i].UniqueId + "_out.msg", SaveOptions.DefaultMsgUnicode);
            }
            // ExEnd:SavingMessagesFromIMAPServer
        }
    }
}