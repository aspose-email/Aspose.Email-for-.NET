using Aspose.Email.Mime;
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

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SavingMessageInDraftStatus
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            // ExStart:SavingMessageInDraftStatus
            // Change properties of an existing MSG file
            string strExistingMsg = @"message.msg";
            
            // Load the existing file in MailMessage and Change the properties
            MailMessage msg = MailMessage.Load(dataDir + strExistingMsg, new MsgLoadOptions());
            msg.Subject += "NEW SUBJECT (updated by Aspose.Email)";
            msg.HtmlBody += "NEW BODY (udpated by Aspose.Email)";
            
            // Create an instance of type MapiMessage from MailMessage, Set message flag to un-sent (draft status) and Save it
            MapiMessage mapiMsg = MapiMessage.FromMailMessage(msg);
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT);
            mapiMsg.Save(dataDir + "SavingMessageInDraftStatus_out.msg");
            // ExEnd:SavingMessageInDraftStatus
        }
    }
}
