using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
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
    class ReadVoteResultsInformation
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // ExStart:ReadVoteResultsInformation
            MapiMessage msg = MapiMessage.FromFile(dataDir + @"AddVotingButtonToExistingMessage.msg");
            foreach (MapiRecipient recipient in msg.Recipients)
            {
                Console.WriteLine(string.Format("Recipient: {0}", recipient.DisplayName));

                // Get the PR_RECIPIENT_AUTORESPONSE_PROP_RESPONSE property
                Console.WriteLine(string.Format("Response: {0}", recipient.Properties[MapiPropertyTag.PR_RECIPIENT_AUTORESPONSE_PROP_RESPONSE].GetString()));

                // Get the PR_RECIPIENT_TRACKSTATUS_TIME property
                Console.WriteLine(string.Format("Response time: {0}", recipient.Properties[MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME].GetDateTime()));

                Console.WriteLine();
            }
            // ExEnd:ReadVoteResultsInformation
        }
    }
}