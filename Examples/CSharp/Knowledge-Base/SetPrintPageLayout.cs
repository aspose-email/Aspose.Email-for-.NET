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
    class SetPrintPageLayout
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();

            // Declare message as an MailMessage instance
            MailMessage message = new MailMessage();
            message.From = "user1@domain.com";
            message.To = "user2@domain.com";
            message.Subject = "My First Mail";
            message.Date = DateTime.Now;
            
            // Body of the email message
            message.IsBodyHtml = true;
            message.HtmlBody = "<html><body><h1>Hello this is html body in Heading 1 format </h1></body></html>";

            // ExStart:SetPrintPageLayout

            // Instantiate an instance of MailPrinter
            var printer = new Aspose.Email.Printing.MailPrinter();

            // Just for testing, get the default property values
            double width = printer.PageWidth;
            double height = printer.PageHeight;

            // Set page layout for printing
            printer.PageUnit = Aspose.Email.Printing.PrinterUnit.Cm;
            printer.MarginTop = 2;
            printer.MarginBottom = 2;
            printer.MarginLeft = 2;
            printer.MarginRight = 2;
            printer.PageWidth = 8;
            printer.PageHeight = 20;
            // ExEnd:SetPrintPageLayout

            // Print the email to an XPS and TIFF file
            printer.Print(message, dataDir + "PrintPageLayoutXPS_out.xps", Aspose.Email.Printing.PrintFormat.XPS);
            printer.Print(message, dataDir + "PrintPageLayoutTIFF_out.tiff", Aspose.Email.Printing.PrintFormat.Tiff);            

        }
    }
}
