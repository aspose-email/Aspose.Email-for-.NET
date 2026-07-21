// Demonstrates how to read contacts from an Outlook PST file and save each one
// to disk in vCard (VCF) format.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveContactInformation
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Contacts");

            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"Outlook.pst"))
            {
                var folderInfo = personalStorage.RootFolder.GetSubFolder("Contacts");
                var saved = 0;

                foreach (var messageInfo in folderInfo.GetContents())
                {
                    var contact = (MapiContact)personalStorage.ExtractMessage(messageInfo).ToMapiMessageItem();
                    var displayName = contact.NameInfo.DisplayName;

                    Console.WriteLine($"Name: {displayName} - {messageInfo.EntryIdString}");

                    contact.Save(outputDir/(displayName + ".vcf"), ContactSaveFormat.VCard);
                    saved++;
                }

                Console.WriteLine($"\nSaved {saved} contact(s) as vCard to {outputDir}");
            }
        }
    }
}
