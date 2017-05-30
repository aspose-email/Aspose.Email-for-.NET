using System;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email
{
    class EmbeddedObjects
    {
        public static void Run()
        {
            // ExStart:EmbeddedObjects
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "EmbeddedImage.msg";

            // Create an instance of the MailMessage class and Set the addresses and Set the content
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("test001@gmail.com");
            mail.To.Add("test001@gmail.com");
            mail.Subject = "This is an email";

            // Create the plain text part It is viewable by those clients that don't support HTML
            AlternateView plainView = AlternateView.CreateAlternateViewFromString("This is my plain text content", null, "text/plain");

            /* Create the HTML part.To embed images, we need to use the prefix 'cid' in the img src value.
            The cid value will map to the Content-Id of a Linked resource. Thus <img src='cid:barcode'> will map to a LinkedResource with a ContentId of //'barcode'. */
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString("Here is an embedded image.<img src=cid:barcode>", null, "text/html");

            // Create the LinkedResource (embedded image) and Add the LinkedResource to the appropriate view
            LinkedResource barcode = new LinkedResource(dataDir + "1.jpg", MediaTypeNames.Image.Jpeg)
            {
                ContentId = "barcode"
            };
            mail.LinkedResources.Add(barcode);
            mail.AlternateViews.Add(plainView);
            mail.AlternateViews.Add(htmlView);
            mail.Save(dataDir + "EmbeddedImage_out.msg", SaveOptions.DefaultMsgUnicode);
            // ExEnd:EmbeddedObjects
            Console.WriteLine(Environment.NewLine + "Message saved with embedded objects successfully at " + dstEmail);
        }
    }
}