using Aspose.Email.Mapi;
using Aspose.Email.PersonalInfo.VCard;
using System;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using https://forum.aspose.com/c/email            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreateV30Contact
    {
        public static void Run()
        {
            // ExStart:1
            // Load the Outlook file
            string dataDir = RunExamples.GetDataDir_Outlook();

            MapiContact contact = new MapiContact();
            contact.NameInfo = new MapiContactNamePropertySet("Jane", "A.", "Buell");
            contact.ProfessionalInfo = new MapiContactProfessionalPropertySet("Aspose Pty Ltd", "Social work assistant");
            contact.PersonalInfo.PersonalHomePage = "Aspose.com";
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("test@test.com");
            contact.Telephones.HomeTelephoneNumber = "06605040000";

            VCardSaveOptions opt = new VCardSaveOptions();
            opt.Version = VCardVersion.V30;
            contact.Save(dataDir + "V30.vcf", opt);
            // ExEnd:1

            Console.WriteLine("CreateV30Contact executed successfully.");
        }
    }
}