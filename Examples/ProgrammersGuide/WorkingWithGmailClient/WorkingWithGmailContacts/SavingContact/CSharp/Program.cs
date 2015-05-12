//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Tests;
using Aspose.Email.Google;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;

namespace SavingContactExample
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

            using (IGmailClient client = Aspose.Email.Google.GmailClient.GetInstance(user.ClientId, user.ClientSecret, user.RefreshToken))
            {

                //Contact contact = new Contact();
                Contact[] contacts = client.GetAllContacts();
                Contact contact = contacts[0];
                contact.Save("contact.msg", ContactSaveFormat.Msg);
                contact.Save("contact.vcf", ContactSaveFormat.VCard);


            }
        }
    }

}