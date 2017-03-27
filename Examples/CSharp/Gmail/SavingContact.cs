using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email.PersonalInfo;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Gmail
{
    class SavingContact
    {
        public static void Run()
        {
            try
            {
                string dataDir = RunExamples.GetDataDir_Gmail();
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Get IGmailclient
                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    Contact[] contacts = client.GetAllContacts();
                    Contact contact = contacts[0];

                    // ExStart:SavingContact
                    contact.Save(dataDir + "contact_out.msg", ContactSaveFormat.Msg);
                    contact.Save(dataDir + "contact_out.vcf", ContactSaveFormat.VCard);
                    // ExEnd:SavingContact
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
