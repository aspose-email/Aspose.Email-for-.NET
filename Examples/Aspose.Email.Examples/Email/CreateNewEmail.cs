// Demonstrates how to create a new MailMessage with an HTML body and save it
// in EML, MSG, and MHTML formats.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class CreateNewEmail
    {
        public static void Run()
        {
            var message = new MailMessage
            {
                From = "from@domain.com",
                To = "to1@domain.com, to2@domain.com",
                CC = "cc1@domain.com, cc2@domain.com",
                Subject = "New message",
                HtmlBody = @"<!DOCTYPE html>
                <html>
                 <head>
                  <style>
                   h3{font-family:Verdana, sans-serif;color:#000000;background-color:#ffffff;}
                   p {font-family:Verdana, sans-serif;font-size:14px;font-style:normal;
                     font-weight:normal;color:#000000;background-color:#ffffff;}
                  </style>
                 </head>
                 <body>
                   <h3>New message</h3>
                   <p>This is a new message created by Aspose.Email.</p>
                 </body>
                </html>"
            };

            // The same message can be written out in any of the supported formats.
            message.Save(Data.Out/"CreateNewEmail_out.eml", SaveOptions.DefaultEml);
            message.Save(Data.Out/"CreateNewEmail_out.msg", SaveOptions.DefaultMsgUnicode);
            message.Save(Data.Out/"CreateNewEmail_out.mhtml", SaveOptions.DefaultMhtml);

            Console.WriteLine($"Created: {message.Subject}");
            Console.WriteLine($"Saved as EML, MSG and MHTML in {Data.Out}");
        }
    }
}
