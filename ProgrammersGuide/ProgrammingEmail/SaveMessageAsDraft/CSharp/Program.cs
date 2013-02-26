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
using Aspose.Email.Outlook;

namespace SaveMessageAsDraft
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create a new instance of MailMessage class
            MailMessage message = new MailMessage();

            // Set sender information
            message.From = "from@domain.com";

            // Add recipients
            message.To.Add("to1@domain.com");
            message.To.Add("to2@domain.com");

            // Set subject of the message
            message.Subject = "New message created by Aspose.Email";

            // Set Html body of the message
            message.IsBodyHtml = true;
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/><font color=blue>This line is in blue color</font>";

            // Create an instance of MapiMessage and load the MailMessag instance into it
            MapiMessage mapiMsg = MapiMessage.FromMailMessage(message);

            // Set the MapiMessageFlags as UNSENT and FROMME
            mapiMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_UNSENT | MapiMessageFlags.MSGFLAG_FROMME);

            // Save the MapiMessage to disk
            mapiMsg.Save(dataDir + "New-Draft.msg");

            // Display Status.
            System.Console.WriteLine("Message saved as draft successfully.");
        }
    }
}