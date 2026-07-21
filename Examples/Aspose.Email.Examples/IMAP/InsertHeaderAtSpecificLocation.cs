using Aspose.Email.Mime;

namespace Aspose.Email.Examples.IMAP
{
    class InsertHeaderAtSpecificLocation
    {
        public static void Run()
        {
            // The path to the File directory.
            string loadFile = Data.Imap + "InsertHeaders.eml";
            MailMessage eml = MailMessage.Load(loadFile);
            eml.Headers.Insert("secret-header", "mystery1");
            eml.Save(Data.Out + "Updated-MessageHeaders_out.eml");
        }
    }
}
