// Demonstrates searching for people with the task-based EWS client.
//
// FindPeopleAsync takes an EwsFindPeople parameter object: SetQueryString searches
// contacts and the directory by name or address, SetMailQuery restricts the search to a
// contacts folder with conditions instead, and SetMaxNumberOfItems caps the results.
// FindPeople needs Exchange 2013 or later.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients.Exchange.WebService.Models;

namespace Aspose.Email.Examples.EWS
{
    internal static class FindPeopleEwsAsync
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            const string searchText = "smith";

            using (var client = await ClientBuilder.EwsAsync(AuthType.ModernWithAppPermission))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30)))
            {
                var people = (await client.FindPeopleAsync(EwsFindPeople.Create()
                    .SetQueryString(searchText)
                    .SetMaxNumberOfItems(10)
                    .SetCancellationToken(cancellation.Token))).ToList();

                Console.WriteLine($"People matching '{searchText}': {people.Count}");
                foreach (var person in people)
                    Console.WriteLine($"  {person.DisplayName} <{person.EmailAddresses.FirstOrDefault()?.Address}>");
            }
        }
    }
}
