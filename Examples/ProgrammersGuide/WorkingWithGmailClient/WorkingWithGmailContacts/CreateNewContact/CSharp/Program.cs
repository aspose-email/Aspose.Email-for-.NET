//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Exchange.Schema;
using Aspose.Email.Tests;
using Aspose.Email.Google;
using Aspose.Email.Mail;
using System;

namespace CreateNewContactExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            string username = "aspose.examples";
            string email = "aspose.examples@gmail.com";
            string password = "aspose123";
            string clientId = "491198589534-c08fdgnpu3ausn9pbdjjjh9r2s4vlla0.apps.googleusercontent.com";
            string clientSecret = "Km9G7a6PivZv4dDxZ3-ZVPk2";
            string refresh_token = "1/UlvQ1pvWN5DQLgHd8LmcNchPs50s7E96sTnjdwHKVS8";

            //The refresh_token is to be used in below.
            GoogleTestUser user = new GoogleTestUser(
                            username,
                            email,
                            password,
                            clientId,     //client id
                            clientSecret,     //client secret
                            refresh_token);   //refresh token

            //Gmail Client
            IGmailClient client = Aspose.Email.Google.GmailClient.GetInstance(user.ClientId, user.ClientSecret, user.RefreshToken, user.EMail);

            //Create a Contact
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

            //Email Address
            Aspose.Email.Mail.EmailAddress eAddress = new Aspose.Email.Mail.EmailAddress();
            eAddress.Address = "email@gmail.com";
            contact.EmailAddresses.Add(eAddress);

            string contactUri = client.CreateContact(contact);
            
        }
    }
}