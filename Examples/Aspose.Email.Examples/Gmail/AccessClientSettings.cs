using Aspose.Email.Clients.Google;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.Gmail
{
    class AccessClientSettings
    {
        public static void Run()
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
