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
    class ManageAccessRule
    {
        public static void Run()
        {
            try
            {
                // ExStart:ManageAccessRule
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    // Retrieve list of calendars for the current client
                    ExtendedCalendar[] calendarList = client.ListCalendars();

                    // Get first calendar id and retrieve list of AccessControlRule for the first calendar
                    string calendarId = calendarList[0].Id;
                    AccessControlRule[] roles1 = client.ListAccessRules(calendarId);

                    // Create a local access control rule and Set rule properties
                    AccessControlRule rule = new AccessControlRule();
                    rule.Role = AccessRole.reader;
                    rule.Scope = new AclScope(AclScopeType.user, User2.EMail);

                    // Insert new rule for the calendar. It returns the newly created rule
                    AccessControlRule createdRule = client.CreateAccessRule(calendarId, rule);

                    // Confirm if local created rule and returned rule are equal
                    if ((rule.Role == createdRule.Role) && (rule.Scope.Type == createdRule.Scope.Type) && (rule.Scope.Value.ToLower() == createdRule.Scope.Value.ToLower()))
                    {
                        Console.WriteLine("local rule and returned rule after creation are equal");
                    }
                    else
                    {
                        Console.WriteLine("Rule could not be created successfully");
                        return;
                    }

                    // Get list of rules
                    AccessControlRule[] roles2 = client.ListAccessRules(calendarId);

                    // Current list length should be 1 more than the earlier one
                    if (roles1.Length + 1 == roles2.Length)
                    {
                        Console.WriteLine("List lengths are ok");
                    }
                    else
                    {
                        Console.WriteLine("List lengths are not ok");
                        return;
                    }

                    // Change rule and Update the rule for the selected calendar
                    createdRule.Role = AccessRole.writer;
                    AccessControlRule updatedRule = client.UpdateAccessRule(calendarId, createdRule);

                    // Check if returned access control rule after update is ok
                    if ((createdRule.Role == updatedRule.Role) && (createdRule.Id == updatedRule.Id))
                    {
                        Console.WriteLine("Rule is updated successfully");
                    }
                    else
                    {
                        Console.WriteLine("Rule is not updated");
                        return;
                    }

                    // Retrieve individaul rule against a calendar
                    AccessControlRule fetchedRule = client.FetchAccessRule(calendarId, createdRule.Id);

                    //Check if rule parameters are ok
                    if ((updatedRule.Id == fetchedRule.Id) && (updatedRule.Role == fetchedRule.Role) && (updatedRule.Scope.Type == fetchedRule.Scope.Type) && (updatedRule.Scope.Value.ToLower() == fetchedRule.Scope.Value.ToLower()))
                    {
                        Console.WriteLine("Rule parameters are ok");
                    }
                    else
                    {
                        Console.WriteLine("Rule parameters are not ok");
                    }

                    // Delete particular rule against a given calendar and Retrieve the all rules list for the same calendar
                    client.DeleteAccessRule(calendarId, createdRule.Id);
                    AccessControlRule[] roles3 = client.ListAccessRules(calendarId);

                    // Check that current rules list length should be equal to the original list length before adding and deleting the rule
                    if (roles1.Length == roles3.Length)
                    {
                        Console.WriteLine("List lengths are same");
                    }
                    else
                    {
                        Console.WriteLine("List lengths are not equal");
                        return;
                    }
                }
                // ExEnd:ManageAccessRule
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
