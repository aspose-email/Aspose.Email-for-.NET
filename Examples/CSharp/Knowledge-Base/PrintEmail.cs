using System.IO;
using System;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class PrintEmail
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            string dstXPS = dataDir + "email-print.xps";
            string dstTIFF = dataDir + "email-print.tiff";

            // Declare message as an MailMessage instance
            MailMessage message = new MailMessage();

            // Sender's address
            message.From = "user1@domain.com";

            // Recipient's address
            message.To = "user2@domain.com";

            // Subject of the email
            message.Subject = "My First Mail";

            // Set message date to current date and time
            message.Date = DateTime.Now;

            // Body of the email message
            message.Body = "Text is the Mail Message";

            // Body of the email message
            message.IsBodyHtml = true;
            message.HtmlBody = "<html><body><h1>Hello this is html body in Heading 1 format </h1></body></html>";

            // Instantiate an instance of MailPrinter
            Aspose.Email.Printing.MailPrinter printer = new Aspose.Email.Printing.MailPrinter();

            // Set the MessageFormattingFlags to none to display only the message body
            printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.None;

            // Set MessageFormattingFlags to MailInfo to display the message headers and body
            printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.MailInfo;

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

            // Print the email to an XPS file
            printer.Print(message, dstXPS, Aspose.Email.Printing.PrintFormat.XPS);

            // Auto-Fit a TIFF
            printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.AutoFitWidth;

            // Print the email to a TIFF file
            printer.Print(message, dstTIFF, Aspose.Email.Printing.PrintFormat.Tiff);

            Console.WriteLine(Environment.NewLine + "Printed email successfully. saved at " + dstXPS);
            Console.WriteLine(Environment.NewLine + "Printed email successfully. saved at " + dstTIFF);
        }
    }
}
