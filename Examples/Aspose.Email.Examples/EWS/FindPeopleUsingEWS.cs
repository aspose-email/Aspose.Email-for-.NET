// Demonstrates searching for people across the user's contacts and the directory.
//
// FindPeople (Exchange 2013 and later) looks up people by name or e-mail address the way
// the Outlook "Search People" box does, covering personal contacts and the global
// address list in one call. The second argument caps the number of results.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class FindPeopleUsingEWS
    {
        public static void Run()
        {
            const string searchText = "smith";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var people = client.FindPeople(searchText, 10);

                Console.WriteLine($"People matching '{searchText}': {people.Length}");
                foreach (var person in people)
                {
                    var email = person.EmailAddresses.FirstOrDefault()?.Address;
                    Console.WriteLine($"  {person.DisplayName} <{email}>");
                }
            }
        }
    }
}
