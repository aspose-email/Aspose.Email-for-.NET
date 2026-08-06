// Demonstrates the strict container-class check when adding a PST subfolder: by
// default a mismatching class is silently accepted, and the option makes it throw.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class EnforceContainerClassMatching
    {
        public static void Run()
        {
            var pstPath = Data.Out/"EnforceContainerClassMatching_out.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                var contacts = pst.CreatePredefinedFolder("Contacts", StandardIpmFolder.Contacts);
                Console.WriteLine($"Parent folder class: {contacts.ContainerClass}");

                // Without the check a note folder can be created under Contacts.
                var lenient = contacts.AddSubFolder("Subfolder1", "IPF.Note");
                Console.WriteLine($"Added without check:  {lenient.DisplayName} ({lenient.ContainerClass})");

                // With EnforceContainerClassMatching the mismatch is rejected instead.
                var options = new FolderCreationOptions
                {
                    EnforceContainerClassMatching = true,
                    ContainerClass = "IPF.Note"
                };

                try
                {
                    contacts.AddSubFolder("Subfolder2", options);
                    Console.WriteLine("Added with check:     unexpectedly succeeded");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Added with check:     rejected - {ex.Message}");
                }
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}
