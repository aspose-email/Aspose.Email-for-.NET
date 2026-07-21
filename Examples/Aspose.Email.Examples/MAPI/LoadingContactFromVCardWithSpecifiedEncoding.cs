// Demonstrates how to load a MapiContact from a VCard file with an explicit text encoding.

using System;
using System.Text;
using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;

namespace Aspose.Email.Examples.MAPI
{
    internal static class LoadingContactFromVCardWithSpecifiedEncoding
    {
        public static void Run()
        {
            // PreferredEncoding is used for vCards that carry no encoding information of
            // their own; the FromVCard(path, Encoding) overload is obsolete.
            var options = new VCardLoadOptions { PreferredEncoding = Encoding.UTF8 };

            var contact = MapiContact.FromVCard(Data.Mapi/"Contact.vcf", options);
            Console.WriteLine($"Name: {contact.NameInfo.DisplayName}");
        }
    }
}
