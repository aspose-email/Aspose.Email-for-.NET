//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Mail;

namespace CreateNewEmail
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create a new instance of MailMessage class
            MailMessage message = new MailMessage();

            // Set subject of the message
            message.Subject = "New message created by Aspose.Email for .NET";

            // Set Html body
            message.IsBodyHtml = true;
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>";

            // Set sender information
            message.From = "from@domain.com";

            // Add TO recipients
            message.To.Add("to1@domain.com");
            message.To.Add("to2@domain.com");

            //Add CC recipients
            message.CC.Add("cc1@domain.com");
            message.CC.Add("cc2@domain.com");

            // Save message in EML, MSG and MHTML formats
            message.Save(dataDir + "Message.eml", MailMessageSaveType.EmlFormat);
            message.Save(dataDir + "Message.msg", MailMessageSaveType.OutlookMessageFormat);
            message.Save(dataDir + "Message.mhtml", MailMessageSaveType.MHtmlFromat);

            // Display Status.
            System.Console.WriteLine("New Emails created successfully.");
        }
    }
}