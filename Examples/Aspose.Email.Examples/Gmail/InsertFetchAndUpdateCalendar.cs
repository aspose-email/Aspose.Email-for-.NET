using Aspose.Email.Clients.Google;
using System;

namespace Aspose.Email.Examples.Gmail
{
    class InsertFetchAndUpdateCalendar
    {
        public static void Run()
        {
            try
            {
                // Get access token
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    // Insert, get and update calendar
                    Aspose.Email.Clients.Google.Calendar calendar = new Aspose.Email.Clients.Google.Calendar("summary - " + Guid.NewGuid().ToString(), null, null, "America/Los_Angeles");
                    
                    // Insert calendar and Retrieve same calendar using id
                    string id = client.CreateCalendar(calendar);
                    Aspose.Email.Clients.Google.Calendar cal = client.FetchCalendar(id);

                    //Match the retrieved calendar info with local calendar
                    if ((calendar.Summary == cal.Summary) && (calendar.TimeZone == cal.TimeZone))
                    {
                        Console.WriteLine("fetched calendar information matches");
                    }
                    else
                    {
                        Console.WriteLine("fetched calendar information does not match");
                    }

                    // Change information in the fetched calendar and Update calendar
                    cal.Description = "Description - " + Guid.NewGuid().ToString();
                    cal.Location = "Location - " + Guid.NewGuid().ToString();
                    client.UpdateCalendar(cal);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
