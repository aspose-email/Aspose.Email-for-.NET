using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Email
{
    class RemoveLRTracesFromMessageBody
    {
        public static void Run()
        {
            // ExStart:CreateNewMailMessage
            string dataDir = RunExamples.GetDataDir_Email();

            //sample input file
            string fileName = "EmlWithLinkedResources.eml";

            //Load the test message with Linked Resources
            MailMessage msg = MailMessage.Load(dataDir + fileName);

            //Remove a LinkedResource
            msg.LinkedResources.RemoveAt(0, true);

            //Now clear the Alternate View for linked Resources
            msg.AlternateViews[0].LinkedResources.Clear(true);
            // ExEnd:CreateNewMailMessage
        }
    }
}
