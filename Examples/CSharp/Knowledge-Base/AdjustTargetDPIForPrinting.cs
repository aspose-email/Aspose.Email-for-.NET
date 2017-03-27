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
    class AdjustTargetDPIForPrinting
    {
        public static void Run()
        {
            // ExStart:AdjustTargetDPIForPrinting
            // The path to the File directory and Instantiate an instance of MailPrinter
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            Printing.MailPrinter printer = new Printing.MailPrinter();
            printer.DpiX = 300;
            printer.DpiY = 300;
            MailMessage msg = MailMessage.Load(dataDir + "message3.msg", new MsgLoadOptions());
            printer.Print(msg, dataDir + "AdjustTargetDPI_out.tiff", Printing.PrintFormat.Tiff);
            // ExEnd:AdjustTargetDPIForPrinting
        }
    }
}
