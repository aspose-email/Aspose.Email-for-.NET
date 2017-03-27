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
    class ReadingNamedMAPIPropertyFromAttachment
    {
        public static void Run()
        {
            string getProperty = GetNamedPropertyByAspose();
            Console.WriteLine("getProperty" + getProperty);
        }

        // Gets attachment's named property
        // Property name is "CustomAttGuid"
        private static string GetNamedPropertyByAspose()
        {
            //ExStart:ReadingNamedMAPIPropertyFromAttachment
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load from file
            MailMessage mail = MailMessage.Load(dataDir + "outputAttachments.msg");

            var mapi = MapiMessage.FromMailMessage(mail);

            foreach (MapiNamedProperty namedProperty in mapi.Attachments[0].NamedProperties.Values)
            {
                if (string.Compare(namedProperty.NameId, "CustomAttGuid", StringComparison.OrdinalIgnoreCase) == 0)
                {

                    return namedProperty.GetString();
                }
            }
            return string.Empty;
            //ExEnd:ReadingNamedMAPIPropertyFromAttachment
        }
    }
}