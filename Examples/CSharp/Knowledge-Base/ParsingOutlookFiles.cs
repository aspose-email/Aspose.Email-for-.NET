using Aspose.Email.Mapi;
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
    class ParsingOutlookFiles
    {
        public static void Run()
        {
            // ExStart:ParsingOutlookFiles
            // The path to the File directory and Load Microsoft Outlook email message file
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            MapiMessage msg = MapiMessage.FromFile(dataDir + "message3.msg");

            // Obtain subject of the email message, sender, body and Attachment count
            Console.WriteLine("Subject:" + msg.Subject);
            Console.WriteLine("From:" + msg.SenderName);
            Console.WriteLine("Body:" + msg.Body);
            Console.WriteLine("Attachment Count:" + msg.Attachments.Count);

            // Iterate through the attachments
            foreach (MapiAttachment attachment in msg.Attachments)
            {
                // Access the attachment's file name and Save attachment
                Console.WriteLine("Attachment:" + attachment.FileName);
                attachment.Save(dataDir + attachment.LongFileName);
            }
            // ExEnd:ParsingOutlookFiles
        }
    }
}
