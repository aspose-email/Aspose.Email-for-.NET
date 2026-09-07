// Demonstrates the structured parts of an Outlook contact that are easy to miss:
// the three physical addresses, which one is flagged as the mailing address, and the
// two dates Outlook tracks separately from the rest of the personal information.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetContactEventsAndAddresses
    {
        public static void Run()
        {
            var contact = new MapiContact
            {
                NameInfo = new MapiContactNamePropertySet("Bertha", "A.", "Buell")
            };
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("bertha@example.com");

            // Birthday and anniversary live in their own property set, not in PersonalInfo.
            contact.Events = new MapiContactEventPropertySet
            {
                Birthday = new DateTime(1980, 3, 14),
                WeddingAnniversary = new DateTime(2005, 9, 27)
            };

            // A contact holds three addresses; each one is a structured value rather
            // than a single string.
            contact.PhysicalAddresses.WorkAddress = new MapiContactPhysicalAddress
            {
                Street = "Im Astenfeld 59",
                City = "Edelschrott",
                StateOrProvince = "Styria",
                PostalCode = "8580",
                Country = "Austria",
                CountryCode = "AT",
                PostOfficeBox = "PO Box 12",

                // Exactly one address should carry this flag - it is the one Outlook
                // uses when it addresses an envelope.
                IsMailingAddress = true
            };

            contact.PhysicalAddresses.HomeAddress = new MapiContactPhysicalAddress
            {
                Street = "Hauptstrasse 4",
                City = "Graz",
                PostalCode = "8010",
                Country = "Austria"
            };

            contact.PhysicalAddresses.OtherAddress = new MapiContactPhysicalAddress
            {
                Street = "Seeweg 1",
                City = "Klagenfurt",
                Country = "Austria"
            };

            var outputPath = Data.Out/"SetContactEventsAndAddresses_out.msg";
            contact.Save(outputPath, ContactSaveFormat.Msg);

            // Read it back so the output reflects what actually persisted.
            var reloaded = (MapiContact)MapiMessage.Load(outputPath).ToMapiMessageItem();

            Console.WriteLine($"Contact: {reloaded.NameInfo.DisplayName}");
            Console.WriteLine($"Birthday:    {Show(reloaded.Events.Birthday)}");
            Console.WriteLine($"Anniversary: {Show(reloaded.Events.WeddingAnniversary)}");

            Print("Work", reloaded.PhysicalAddresses.WorkAddress);
            Print("Home", reloaded.PhysicalAddresses.HomeAddress);
            Print("Other", reloaded.PhysicalAddresses.OtherAddress);

            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static void Print(string label, MapiContactPhysicalAddress address)
        {
            if (address == null)
            {
                Console.WriteLine($"\n{label} address: (none)");
                return;
            }

            Console.WriteLine($"\n{label} address (mailing: {address.IsMailingAddress}):");
            Console.WriteLine($"  street:   {Show(address.Street)}");
            Console.WriteLine($"  city:     {Show(address.City)}");
            Console.WriteLine($"  region:   {Show(address.StateOrProvince)}");
            Console.WriteLine($"  postcode: {Show(address.PostalCode)}");
            Console.WriteLine($"  country:  {Show(address.Country)} ({Show(address.CountryCode)})");
            Console.WriteLine($"  PO box:   {Show(address.PostOfficeBox)}");

            // Address is the whole thing as one preformatted block.
            Console.WriteLine($"  formatted: {Show(address.Address).Replace("\r\n", " / ")}");
        }

        private static string Show(string value)
        {
            return string.IsNullOrEmpty(value) ? "(not set)" : value;
        }

        private static string Show(DateTime value)
        {
            return value == DateTime.MinValue ? "(not set)" : value.ToString("yyyy-MM-dd");
        }
    }
}
