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
using Aspose.Email.Mime;

namespace EmbeddedObjects
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Create an instance of the MailMessage class
            MailMessage mail = new MailMessage();

            //Set the addresses
            mail.From = new MailAddress("test001@kerio.com");
            mail.To.Add("test001@kerio.com");

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
            
            mail.Save(dataDir + "EmbeddedImage.msg", MessageFormat.Msg);
        }
    }
}