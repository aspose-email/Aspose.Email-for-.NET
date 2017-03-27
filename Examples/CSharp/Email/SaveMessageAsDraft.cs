using System.IO;
using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Pop3;
using Aspose.Email;
using Aspose.Email.Clients.Imap;
using System.Configuration;
using System.Data;

namespace Aspose.Email.Examples.CSharp.Email
{
    class SaveMessageAsDraft
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "New-Draft.msg";

            // Create a new instance of MailMessage class
            MailMessage message = new MailMessage();

            // Set sender information
            message.From = "from@domain.com";

            // Add recipients
            message.To.Add("to1@domain.com");
            message.To.Add("to2@domain.com");

            // Set subject of the message
            message.Subject = "New message created by Aspose.Email";

            // Set Html body of the message
            message.IsBodyHtml = true;
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>";

            // Create an instance of MapiMessage and load the MailMessag instance into it
            MapiMessage mapiMsg = MapiMessage.FromMailMessage(message);

            // Set the MapiMessageFlags as UNSENT and FROMME
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT | MapiMessageFlags.MSGFLAG_FROMME);

            // Save the MapiMessage to disk
            mapiMsg.Save(dstEmail);

            Console.WriteLine(Environment.NewLine + "Created draft MSG at " + dstEmail);
        }
    }
}
