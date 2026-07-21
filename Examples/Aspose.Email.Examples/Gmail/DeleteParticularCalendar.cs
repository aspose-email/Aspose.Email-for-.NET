using Aspose.Email.Clients.Google;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.Gmail
{
    class DeleteParticularCalendar
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
                    // Access and delete calendar with summary starting from "Calendar summary - "
                    string summary = "Calendar summary - ";

                    // Get calendars list
                    ExtendedCalendar[] lst0 = client.ListCalendars();

                    foreach (ExtendedCalendar extCal in lst0)
                    {
                        // Delete selected calendars
                        if (extCal.Summary.StartsWith(summary))
                            client.DeleteCalendar(extCal.Id);
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
