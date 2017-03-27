using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Gmail
{
    class CreateGmailContact
    {
        public static void Run()
        {
            try
            {
                // ExStart:CreateGmailContact
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Gmail Client
                IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail);

                // Create a Contact
                Contact contact = new Contact();
                contact.Prefix = "Prefix";
                contact.GivenName = "GivenName";
                contact.Surname = "Surname";
                contact.MiddleName = "MiddleName";
                contact.DisplayName = "Test User 1";
                contact.Suffix = "Suffix";
                contact.JobTitle = "JobTitle";
                contact.DepartmentName = "DepartmentName";
                contact.CompanyName = "CompanyName";
                contact.Profession = "Profession";
                contact.Notes = "Notes";
                PostalAddress address = new PostalAddress();
                address.Category = PostalAddressCategory.Work;
                address.Address = "Address";
                address.Street = "Street";
                address.PostOfficeBox = "PostOfficeBox";
                address.City = "City";
                address.StateOrProvince = "StateOrProvince";
                address.PostalCode = "PostalCode";
                address.Country = "Country";
                contact.PhysicalAddresses.Add(address);
                PhoneNumber pnWork = new PhoneNumber();
                pnWork.Number = "323423423423";
                pnWork.Category = PhoneNumberCategory.Work;
                contact.PhoneNumbers.Add(pnWork);
                PhoneNumber pnHome = new PhoneNumber();
                pnHome.Number = "323423423423";
                pnHome.Category = PhoneNumberCategory.Home;
                contact.PhoneNumbers.Add(pnHome);
                PhoneNumber pnMobile = new PhoneNumber();
                pnMobile.Number = "323423423423";
                pnMobile.Category = PhoneNumberCategory.Mobile;
                contact.PhoneNumbers.Add(pnMobile);
                contact.Urls.Blog = "Blog.ru";
                contact.Urls.BusinessHomePage = "BusinessHomePage.ru";
                contact.Urls.HomePage = "HomePage.ru";
                contact.Urls.Profile = "Profile.ru";
                contact.Events.Birthday = DateTime.Now.AddYears(-30);
                contact.Events.Anniversary = DateTime.Now.AddYears(-10);
                contact.InstantMessengers.AIM = "AIM";
                contact.InstantMessengers.GoogleTalk = "GoogleTalk";
                contact.InstantMessengers.ICQ = "ICQ";
                contact.InstantMessengers.Jabber = "Jabber";
                contact.InstantMessengers.MSN = "MSN";
                contact.InstantMessengers.QQ = "QQ";
                contact.InstantMessengers.Skype = "Skype";
                contact.InstantMessengers.Yahoo = "Yahoo";
                contact.AssociatedPersons.Spouse = "Spouse";
                contact.AssociatedPersons.Sister = "Sister";
                contact.AssociatedPersons.Relative = "Relative";
                contact.AssociatedPersons.ReferredBy = "ReferredBy";
                contact.AssociatedPersons.Partner = "Partner";
                contact.AssociatedPersons.Parent = "Parent";
                contact.AssociatedPersons.Mother = "Mother";
                contact.AssociatedPersons.Manager = "Manager";

                // Email Address
                EmailAddress eAddress = new EmailAddress();
                eAddress.Address = "email@gmail.com";
                contact.EmailAddresses.Add(eAddress);
                string contactUri = client.CreateContact(contact);
                // ExEnd:CreateGmailContact
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
