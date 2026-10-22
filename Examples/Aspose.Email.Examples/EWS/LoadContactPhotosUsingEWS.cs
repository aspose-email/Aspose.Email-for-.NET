// Demonstrates downloading contact photos only when they are needed.
//
// Listing contacts with ExchangeListContactsOptions.None skips the photo data, which
// keeps the listing fast. Contact.Photo still tells you that a photo exists, and
// LoadContactPhoto fetches its bytes on demand - here for the first few contacts, saved
// to Out.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class LoadContactPhotosUsingEWS
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("EwsContactPhotos");

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var contacts = client.GetContacts(client.MailboxInfo.ContactsUri, ExchangeListContactsOptions.None);
                var withPhoto = contacts.Where(contact => contact.Photo != null).Take(5).ToList();

                Console.WriteLine($"{contacts.Length} contact(s), {withPhoto.Count} with a photo loaded below:");

                foreach (var contact in withPhoto)
                {
                    client.LoadContactPhoto(contact.Photo);

                    var fileName = $"{contact.DisplayName ?? "contact"}.{contact.Photo.PhotoImageFormat}";
                    var safeName = string.Concat(fileName.ToLowerInvariant().Split(Path.GetInvalidFileNameChars()));
                    var outputPath = Path.Combine(outputDir, safeName);
                    File.WriteAllBytes(outputPath, contact.Photo.Data);

                    Console.WriteLine($"  {contact.DisplayName}: {contact.Photo.Data.Length} bytes -> {outputPath}");
                }
            }
        }
    }
}
