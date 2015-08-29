using Aspose.Email.Outlook;

namespace Aspose_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new MapiContact Object
            MapiContact mapiContact = new MapiContact();

            //Set different properties of this Contact object
            mapiContact.NameInfo = new MapiContactNamePropertySet("Mellissa", "", "MacBeth");
            mapiContact.ProfessionalInfo.Title = "Account Representative";
            mapiContact.ProfessionalInfo.CompanyName = "Contoso Ltd.";
            mapiContact.ProfessionalInfo.OfficeLocation = "36/2529";
            mapiContact.Telephones.BusinessTelephoneNumber = "4255551212 x432";
            mapiContact.PhysicalAddresses.WorkAddress.Street = "1 Microsoft Way";
            mapiContact.PhysicalAddresses.WorkAddress.City = "Redmond";
            mapiContact.PhysicalAddresses.WorkAddress.StateOrProvince = "WA";
            mapiContact.PhysicalAddresses.WorkAddress.PostalCode = "98052";
            mapiContact.PhysicalAddresses.WorkAddress.Country = "United States of America";
            mapiContact.ElectronicAddresses.Email1.EmailAddress = "milissa@contoso.com";
            mapiContact.ElectronicAddresses.Email1.AddressType = "SMTP";
            mapiContact.ElectronicAddresses.Email1.DisplayName = "Melissa MacBeth (mellissa@contoso.com)";

            //Save the Contact object to disc
            mapiContact.Save("Contact.vcf", ContactSaveFormat.VCard);
        }
    }
}
