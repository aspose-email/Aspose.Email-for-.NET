using Aspose.Email.Clients.Google;
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
    class AccessClientSettings
    {
        public static void Run()
        {
            try
            {
                // ExStart:AccessClientSettings
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    // Retrieve client settings
                    Dictionary<string, string> settings = client.GetSettings();
                    if (settings.Count < 1)
                    {
                        Console.WriteLine("No settings are available.");
                        return;
                    }

                    // Traverse the settings list
                    foreach (KeyValuePair<string, string> pair in settings)
                    {
                        // Get the setting value and test if settings are ok
                        string value = client.GetSetting(pair.Key);
                        if (pair.Value == value)
                        {
                            Console.WriteLine("Key = " + pair.Key + ", Value = " + pair.Value);
                        }
                        else
                        {
                            Console.WriteLine("Settings could not be retrieved");
                        }
                    }
                }
                // ExEnd:AccessClientSettings
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
