using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class EmbeddedObjects
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "EmbeddedImage.msg";

            //Create an instance of the MailMessage class
            MailMessage mail = new MailMessage();

            //Set the addresses
            mail.From = new MailAddress("test001@gmail.com");
            mail.To.Add("test001@gmail.com");

            //Set the content
            mail.Subject = "This is an email";

            //Create the plain text part
            //It is viewable by those clients that don't support HTML
            AlternateView plainView = AlternateView.CreateAlternateViewFromString("This is my plain text content", null, "text/plain");

            //Create the HTML part.
            //To embed images, we need to use the prefix 'cid' in the img src value.
            //The cid value will map to the Content-Id of a Linked resource.
            //Thus <img src='cid:barcode'> will map to a LinkedResource with a ContentId of //'barcode'.

            AlternateView htmlView = AlternateView.CreateAlternateViewFromString("Here is an embedded image.<img src=cid:barcode>", null, "text/html");

            //create the LinkedResource (embedded image)

            LinkedResource barcode = new LinkedResource(dataDir + "barcode.png", MediaTypeNames.Image.Png);
            barcode.ContentId = "barcode";

            //Add the LinkedResource to the appropriate view

            mail.LinkedResources.Add(barcode);

            mail.AlternateViews.Add(plainView);
            mail.AlternateViews.Add(htmlView);

            mail.Save(dstEmail, Aspose.Email.Mail.SaveOptions.DefaultMsgUnicode);

            Console.WriteLine(Environment.NewLine + "Message saved with embedded objects successfully at " + dstEmail);
        }
    }
}
