using Aspose.Email.Clients.Google;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.Gmail
{
    class AccessColorInfo
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
                    ColorsInfo colors = client.GetColors();
                    Dictionary<string, Colors> palettes = colors.Calendar;

                    // Traverse the settings list
                    foreach (KeyValuePair<string, Colors> pair in palettes)
                    {
                        Console.WriteLine("Key = " + pair.Key + ", Color = " + pair.Value);
                    }
                    Console.WriteLine("Update Date = " + colors.Updated);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
