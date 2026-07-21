using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email.PersonalInfo;

namespace Aspose.Email.Examples.Gmail
{
    class SavingContact
    {
        public static void Run()
        {
            try
            {
                string dataDir = Data.Gmail;
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Get IGmailclient
                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    Contact[] contacts = client.GetAllContacts();
                    Contact contact = contacts[0];

                    contact.Save(dataDir + "contact_out.msg", ContactSaveFormat.Msg);
                    contact.Save(dataDir + "contact_out.vcf", ContactSaveFormat.VCard);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
