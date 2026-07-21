// Demonstrates how to read contacts from an Outlook PST file and save each one as a MSG file.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AccessContactInformation
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Contacts");

            using (var pst = PersonalStorage.FromFile(Data.Mapi/"SampleContacts.pst"))
            {
                var folder = pst.RootFolder.GetSubFolder("Contacts");
                var saved = 0;

                foreach (var info in folder.GetContents())
                {
                    var message = pst.ExtractMessage(info);
                    var contact = (MapiContact)message.ToMapiMessageItem();

                    Console.WriteLine($"Name: {contact.NameInfo.DisplayName}");

                    if (contact.NameInfo.DisplayName != null)
                    {
                        var safeName = message.Subject
                            .Replace(":", " ").Replace("\\", " ")
                            .Replace("?", " ").Replace("/", " ");
                        message.Save(outputDir/(safeName + "_out.msg"));
                        saved++;
                    }
                }

                Console.WriteLine($"\nSaved {saved} contact(s) to {outputDir}");
            }
        }
    }
}
