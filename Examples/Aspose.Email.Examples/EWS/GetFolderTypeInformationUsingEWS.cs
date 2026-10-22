// Demonstrates telling folders apart by what they hold.
//
// FolderType says whether a folder holds mail, appointments, contacts, tasks, notes or
// journal entries, and FolderClass gives the underlying Outlook class (IPF.Note,
// IPF.Appointment, ...). Use them to pick the right API for a folder - ListMessages,
// ListAppointments, GetContacts or ListTasks.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetFolderTypeInformationUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var folder in client.ListSubFolders(client.MailboxInfo.RootUri))
                    Console.WriteLine($"  {folder.DisplayName,-28} {folder.FolderType,-12} {folder.FolderClass}");
            }
        }
    }
}
