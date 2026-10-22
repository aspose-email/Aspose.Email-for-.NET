// Demonstrates working with Outlook notes (sticky notes) in the Notes folder.
//
// CreateItem stores any MAPI item - here a MapiNote - in a folder and returns its URI.
// FetchMapiNotes reads notes back as MapiNote objects, and UpdateNote writes changes to
// an existing note. The example creates a note, changes it, and deletes it again.

using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateAndUpdateNotesUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var note = new MapiNote("Shopping list", "Milk, bread")
                {
                    Color = NoteColor.Yellow
                };

                var uri = client.CreateItem(client.MailboxInfo.NotesUri, note);

                try
                {
                    var stored = client.FetchMapiNotes(new[] { uri })[0];
                    Console.WriteLine($"Created:  {stored.Subject} / {stored.Body.Trim()} / {stored.Color}");

                    stored.Body = "Milk, bread, coffee";
                    stored.Color = NoteColor.Green;
                    client.UpdateNote(uri, stored);

                    var updated = client.FetchMapiNotes(new[] { uri })[0];
                    Console.WriteLine($"Updated:  {updated.Subject} / {updated.Body.Trim()} / {updated.Color}");
                }
                finally
                {
                    client.DeleteItem(uri, DeletionOptions.DeletePermanently);
                }
            }
        }
    }
}
