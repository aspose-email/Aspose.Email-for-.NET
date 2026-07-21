using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.Gmail
{
    class UpdateGmailContact
    {
        public static void Run()
        {
            try
            {
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Get IGmailclient
                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    Contact[] contacts = client.GetAllContacts();
                    Contact contact = contacts[0];
                    contact.JobTitle = "Manager IT";
                    contact.DepartmentName = "Customer Support";
                    contact.CompanyName = "Aspose";
                    contact.Profession = "Software Developer";
                    client.UpdateContact(contact);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
