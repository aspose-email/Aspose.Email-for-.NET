//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Email;
using Aspose.Email.Mail;
using Aspose.Email.Printing;

namespace PrintingFeatures
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Declare message as an MailMessage instance.
            MailMessage message = new MailMessage();

            //Sender's address
            message.From = "user1@domain.com";

            //Recipient's address
            message.To = "user2@domain.com";

            //Subject of the email
            message.Subject = "My First Mail";

            //Set message date to current date and time
            message.Date = DateTime.Now;

            //Body of the email message
            message.TextBody = "Text is the Mail Message";
            
            //Body of the email message
            message.IsBodyHtml = true;
            message.HtmlBody = "<html><body><h1>Hello this is html body in Heading 1 format </h1></body></html>";

            //Instantiate an instance of MailPrinter
            var printer = new Aspose.Email.Printing.MailPrinter();

            //Set the MessageFormattingFlags to none to display only the message body
            printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.None;

            //Set MessageFormattingFlags to MailInfo to display the message headers and body
            printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.MailInfo;

            //Just for testing, get the default property values
            double width = printer.PageWidth;
            double height = printer.PageHeight;

            //Set page layout for printing
            printer.PageUnit = Aspose.Email.Printing.PrinterUnit.Cm;
            printer.MarginTop = 2;
            printer.MarginBottom = 2;
            printer.MarginLeft = 2;
            printer.MarginRight = 2;
            printer.PageWidth = 8;
            printer.PageHeight = 20;

            //Print the email to an XPS file
            printer.Print(message, dataDir + "test.xps", Aspose.Email.Printing.PrintFormat.XPS);

            // Auto-Fit a TIFF
            printer.FormattingFlags = Aspose.Email.Printing.MessageFormattingFlags.AutoFitWidth;

            //Print the email to a TIFF file
            printer.Print(message, dataDir + "test.tiff", Aspose.Email.Printing.PrintFormat.Tiff);

            // Display Status.
            System.Console.WriteLine("Printing to XPS and TIFF completed successfully.");
        }
    }
}