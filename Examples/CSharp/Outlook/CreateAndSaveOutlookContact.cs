using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreateAndSaveOutlookContact
    {
        public static void Run()
        {
            // ExStart:CreateAndSaveOutlookContact
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            MapiContact contact = new MapiContact();
            contact.NameInfo = new MapiContactNamePropertySet("Bertha", "A.", "Buell");
            contact.ProfessionalInfo = new MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant");
            contact.PersonalInfo.PersonalHomePage = "B2BTies.com";
            contact.PhysicalAddresses.WorkAddress.Address = "Im Astenfeld 59 8580 EDELSCHROTT";
            contact.ElectronicAddresses.Email1 = new MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com");
            contact.Telephones = new MapiContactTelephonePropertySet("06605045265");
            contact.PersonalInfo.Children = new string[] { "child1", "child2", "child3" };
            contact.Categories = new string[] { "category1", "category2", "category3" };
            contact.Mileage = "Some test mileage";
            contact.Billing = "Test billing information";
            contact.OtherFields.Journal = true;
            contact.OtherFields.Private = true;
            contact.OtherFields.ReminderTime = new DateTime(2014, 1, 1, 0, 0, 55);
            contact.OtherFields.ReminderTopic = "Test topic";
            contact.OtherFields.UserField1 = "ContactUserField1";
            contact.OtherFields.UserField2 = "ContactUserField2";
            contact.OtherFields.UserField3 = "ContactUserField3";
            contact.OtherFields.UserField4 = "ContactUserField4";

            // Add a photo
            using (FileStream fs = File.OpenRead(dataDir + "Desert.jpg"))
            {
                byte[] buffer = new byte[fs.Length];
                fs.Read(buffer, 0, buffer.Length);
                contact.Photo = new MapiContactPhoto(buffer,
                    MapiContactPhotoImageFormat.Jpeg);
            }
            // Save the Contact in MSG format
            contact.Save(dataDir + "MapiContact_out.msg",ContactSaveFormat.Msg);

            // Save the Contact in VCF format
            contact.Save(dataDir + "MapiContact_out.vcf", ContactSaveFormat.VCard);
            // ExEnd:CreateAndSaveOutlookContact
        }
    }
}