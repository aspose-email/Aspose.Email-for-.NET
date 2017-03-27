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
    class UpdateGmailContact
    {
        public static void Run()
        {
            try
            {
                // ExStart:UpdateGmailContact
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
                // ExEnd:UpdateGmailContact
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
