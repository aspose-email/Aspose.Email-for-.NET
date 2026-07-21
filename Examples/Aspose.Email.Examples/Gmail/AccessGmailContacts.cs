using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.Gmail
{
    class AccessGmailContacts
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
                    foreach (Contact contact in contacts)
                        Console.WriteLine(contact.DisplayName + ", " + contact.EmailAddresses[0]);

                    // Fetch contacts from a specific group
                    ContactGroupCollection groups = client.GetAllGroups();
                    GoogleContactGroup group = null;
                    foreach (GoogleContactGroup g in groups)
                        switch (g.Title)
                        {
                            case "TestGroup": group = g;
                                break;
                        }

                    // Retrieve contacts from the Group
                    if (group != null)
                    {
                        Contact[] contacts2 = client.GetContactsFromGroup(group.Id);
                        foreach (Contact con in contacts2)
                            Console.WriteLine(con.DisplayName + "," + con.EmailAddresses[0].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
