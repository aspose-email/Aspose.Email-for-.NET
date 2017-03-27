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
    class GetMAPIProperties
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();

            // Load mail message
            MapiMessage msg = MapiMessage.FromFile(dataDir + "message3.msg");

            // ExStart:GetMAPIProperties
            MapiProperty mapi = msg.Properties[MapiPropertyTag.PR_SUBJECT_W];
            
            if (mapi.Name.Trim().Length > 0)
            {
                // Display the MAPI property name and value
                Console.WriteLine(mapi.Name);
                Console.WriteLine(mapi.ToString());
            }
            // ExEnd:GetMAPIProperties
        }
    }
}
