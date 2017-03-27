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
    class SetFollowUpForRecipients
    {
        public static void Run()
        {
            // ExStart:SetFollowUpForRecipients
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            MailMessage mailMsg = new MailMessage();
            mailMsg.Sender = "AETest12@gmail.com";
            mailMsg.To = "receiver@gmail.com";
            mailMsg.Body = "This message will test if follow up options can be added to a new mapi message.";

            MapiMessage mapi = MapiMessage.FromMailMessage(mailMsg);
            mapi.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT);  // Mark this message as draft

            DateTime dtReminderDate = new DateTime(2013, 5, 23, 16, 40, 0);

            // Add the follow up flag for receipient now
            FollowUpManager.SetFlagForRecipients(mapi, "Follow up", dtReminderDate);
            mapi.Save(dataDir + "SetFollowUpForRecipients_out.msg");
            // ExEnd:SetFollowUpForRecipients
        }
    }
}