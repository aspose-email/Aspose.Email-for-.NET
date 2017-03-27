using Aspose.Email.Mapi;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
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
    class ReadingOnlyVotingButtons
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();
            string fileName = dataDir + "MessageWithVotingButtons.msg";

            using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(fileName)))
            {
                // ExStart:ReadingOnlyVotingButtons
                MapiMessage testMsg = MapiMessage.FromStream(ms);
                
                // This method can be useful when it is necessary to read only voting buttons Voting buttons will be introduced as a collection of string values
                IList buttons = FollowUpManager.GetVotingButtons(testMsg);
                // ExEnd:ReadingOnlyVotingButtons
            }
            
        }
    }
}
