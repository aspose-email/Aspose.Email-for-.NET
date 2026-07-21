// Demonstrates how to compose an email with an embedded (inline) image using
// AlternateView and LinkedResource, then save it as MSG.

using System;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.Email
{
    internal static class EmbeddedObjects
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "AndrewIrwin@from.com",
                To = "SusanMarc@to.com",
                Subject = "This is an email"
            };

            var plainView = AlternateView.CreateAlternateViewFromString(
                "This is my plain text content", null, "text/plain");

            // Reference the embedded image via cid: scheme in the img src attribute.
            // The cid value must match the ContentId of the LinkedResource.
            var htmlView = AlternateView.CreateAlternateViewFromString(
                "Here is an embedded image.<img src=cid:barcode>", null, "text/html");

            var barcode = new LinkedResource(Data.Email/"1.jpg", MediaTypeNames.Image.Jpeg)
            {
                ContentId = "barcode"
            };

            eml.LinkedResources.Add(barcode);
            eml.AlternateViews.Add(plainView);
            eml.AlternateViews.Add(htmlView);

            eml.Save(Data.Out/"EmbeddedImage_out.msg", SaveOptions.DefaultMsgUnicode);
            Console.WriteLine("Message with embedded image saved.");
        }
    }
}
