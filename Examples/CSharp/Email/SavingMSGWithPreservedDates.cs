using Aspose.Email.Mime;
using System;
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

namespace Aspose.Email.Examples.CSharp.Email
{
    class SavingMSGWithPreservedDates
    {
        public static void Run()
        {
            // ExStart:SavingMSGWithPreservedDates
            // Data directory for reading and writing files
            string dataDir = RunExamples.GetDataDir_Email();

            // Initialize and Load an existing EML file by specifying the MessageFormat
            MailMessage eml = MailMessage.Load(dataDir + "Message.eml");

            // Save as msg with preserved dates
            MsgSaveOptions msgSaveOptions = new MsgSaveOptions(MailMessageSaveType.OutlookMessageFormatUnicode)
            {
                PreserveOriginalDates = true
            };

            eml.Save(Path.Combine(dataDir, "outTest_out.msg"), msgSaveOptions);
            // ExEnd:SavingMSGWithPreservedDates
        }
    }
}
