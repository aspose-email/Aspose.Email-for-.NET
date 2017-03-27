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
    class AccessColorInfo
    {
        public static void Run()
        {
            try
            {
                // ExStart:AccessColorInfo
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
                // ExEnd:AccessColorInfo
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
