using Aspose.Email.Mime;
using Aspose.Email.Storage.Zimbra;
using System;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class SaveMessagesFromZimbraTgzStorage
    {
        public static void Run()
        {
            // ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string outputDir = RunExamples.GetDataDir_Email() + "Zimbra/";

            using (TgzReader reader = new TgzReader(dataDir + "ZimbraSample.tgz"))
            {
                reader.ExportTo(outputDir);
            }
            // ExEnd:1
        }
    }
}
