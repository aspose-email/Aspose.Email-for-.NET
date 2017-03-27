using System;
using Aspose.Email.Mapi;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ParseOutlookMessageFile
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load Microsoft Outlook email message file
            MapiMessage msg = MapiMessage.FromMailMessage(dataDir + "Message.eml");

            // Obtain subject of the email message, SenderName, body and attachments count
            Console.WriteLine("Subject:" + msg.Subject);
            Console.WriteLine("From:" + msg.SenderName);
            Console.WriteLine("Body:" + msg.Body);
            Console.WriteLine("Attachment Count:" + msg.Attachments.Count);

            // Iterate through the attachments
            foreach (MapiAttachment attachment in msg.Attachments)
            {
                Console.WriteLine("Attachment:" + attachment.FileName);
                attachment.Save(attachment.LongFileName);
            }
        }
    }
}