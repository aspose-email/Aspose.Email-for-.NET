using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ReadAndWritingOutlookTemplateFile
    {
        public static void Run()
        {
            //ExStart:ReadAndWritingOutlookTemplateFile
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook template (OFT) file in MailMessage's instance
            MailMessage message = MailMessage.Load(dataDir + "sample.oft", new MsgLoadOptions());

            // Set the sender and recipients information
            string senderDisplayName = "John";
            string senderEmailAddress = "john@abc.com";
            string recipientDisplayName = "William";
            string recipientEmailAddress = "william@xzy.com";

            message.Sender = new MailAddress(senderEmailAddress, senderDisplayName);
            message.To.Add(new MailAddress(recipientEmailAddress, recipientDisplayName));
            message.HtmlBody = message.HtmlBody.Replace("DisplayName", "<b>" + recipientDisplayName + "</b>");

            // Set the name, location and time in email body
            string meetingLocation = "<u>" + "Hall 1, Convention Center, New York, USA" + "</u>";
            string meetingTime = "<u>" + "Monday, June 28, 2010" + "</u>";
            message.HtmlBody = message.HtmlBody.Replace("MeetingPlace", meetingLocation);
            message.HtmlBody = message.HtmlBody.Replace("MeetingTime", meetingTime);

            // Save the message in MSG format and open in Office Outlook
            MapiMessage msg = MapiMessage.FromMailMessage(message);
            msg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT);
            msg.Save(dataDir + "ReadAndWritingOutlookTemplateFile_out.msg");
            //ExEnd:ReadAndWritingOutlookTemplateFile
        }
    }
}