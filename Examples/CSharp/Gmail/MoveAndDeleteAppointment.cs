using Aspose.Email.Calendar;
using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
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
    class MoveAndDeleteAppointment
    {
        public static void Run()
        {
            try
            {
                // ExStart:MoveAndDeleteAppointment
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Get IGmailclient
                using (IGmailClient client = Aspose.Email.Clients.Google.GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    string SourceCalendarId = client.ListCalendars()[0].Id;
                    string DestinationCalendarId = client.ListCalendars()[1].Id;
                    string TargetAppUniqueId = client.ListAppointments(SourceCalendarId)[0].UniqueId;

                    // Retrieve the list of appointments in the destination calendar before moving the appointment
                    Appointment[] appointments = client.ListAppointments(DestinationCalendarId);
                    Console.WriteLine("Before moving count = " + appointments.Length);
                    Appointment Movedapp = client.MoveAppointment(SourceCalendarId, DestinationCalendarId, TargetAppUniqueId);

                    // Retrieve the list of appointments in the destination calendar after moving the appointment
                    appointments = client.ListAppointments(DestinationCalendarId);
                    Console.WriteLine("After moving count = " + appointments.Length);

                    // Delete particular appointment from a calendar using unique id
                    client.DeleteAppointment(DestinationCalendarId, Movedapp.UniqueId);

                    // Retrieve the list of appointments. It should be one less than the earlier appointments in the destination calendar
                    appointments = client.ListAppointments(DestinationCalendarId);
                    Console.WriteLine("After deleting count = " + appointments.Length);
                }

                // ExEnd:MoveAndDeleteAppointment
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
