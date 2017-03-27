using System;
using System.Collections.Generic;
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
    class ReadiNamedMAPIPropertyFromAttachment
    {
        public static void Run()
        {
            //ExStart:ReadiNamedMAPIPropertyFromAttachment
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            
             // Load from file
            MapiMessage msg = MapiMessage.FromFile(dataDir + @"message.msg");

            string subject;

            // Access the MapiPropertyTag.PR_SUBJECT property
            MapiProperty prop = msg.Properties[MapiPropertyTag.PR_SUBJECT];

            // If the property is not found, check the MapiPropertyTag.PR_SUBJECT_W (which is a // Unicode peer of the MapiPropertyTag.PR_SUBJECT)
            if (prop == null)
            {
                prop = msg.Properties[MapiPropertyTag.PR_SUBJECT_W];
            }

            // Cannot found
            if (prop == null)
            {
                Console.WriteLine("No property found!");
                return;
            }

            // Get the property data as string
            subject = prop.GetString();

            Console.WriteLine("Subject:" + subject);
            // Read internet code page property
            prop = msg.Properties[MapiPropertyTag.PR_INTERNET_CPID];
            if (prop != null)
            {
                Console.WriteLine("CodePage:" + prop.GetLong());
            }
            //ExEnd:ReadiNamedMAPIPropertyFromAttachment
        }
    }
}