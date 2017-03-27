using Aspose.Email.Mime;
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

namespace Aspose.Email.Examples.CSharp.Email
{
    class LoadingEMLAndSavingToMSG
    {
        public static void Run()
        {
            // ExStart:LoadingEMLAndSavingToMSG
            // Data directory for reading and writing files
            string dataDir = RunExamples.GetDataDir_Email();

            // Initialize and Load an existing EML file by specifying the MessageFormat
            MailMessage eml = MailMessage.Load(dataDir + "Message.eml");

            // Save the Email message to disk in ASCII format and Unicode format
            eml.Save(dataDir + "AnEmail_out.msg", SaveOptions.DefaultMsgUnicode);                      
            // ExEnd:LoadingEMLAndSavingToMSG
        }
    }
}
