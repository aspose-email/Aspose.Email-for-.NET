// Demonstrates how to add or update standard and custom named MAPI properties on
// every message of a PST folder.

using System;
using System.IO;
using System.Text;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class UpdatePSTCustomProperites
    {
        public static void Run()
        {
            // Work on a copy: the update rewrites messages, and examples must never
            // modify the shared input data.
            var pstPath = Data.Out/"UpdatePSTCustomProperites_out.pst";
            File.Copy(Data.Mapi/"Outlook.pst", pstPath, true);

            using (var personalStorage = PersonalStorage.FromFile(pstPath))
            {
                var testFolder = personalStorage.RootFolder.GetSubFolder("Inbox");

                var newProperties = new MapiPropertyCollection();

                // A standard property, addressed by its tag.
                var property = new MapiProperty(
                    MapiPropertyTag.PR_ORG_EMAIL_ADDR_W,
                    Encoding.Unicode.GetBytes("test_address@org.com"));

                // A custom named property, identified by name within its own GUID.
                var namedProperty1 = new MapiNamedProperty(
                    GenerateNamedPropertyTag(0, MapiPropertyType.PT_LONG),
                    "ITEM_ID", Guid.NewGuid(), BitConverter.GetBytes(123));

                // A known named property - PidLidLogFlags, from the Outlook journal set.
                var namedProperty2 = new MapiNamedProperty(
                    GenerateNamedPropertyTag(1, MapiPropertyType.PT_LONG),
                    0x0000870C, new Guid("0006200A-0000-0000-C000-000000000046"),
                    BitConverter.GetBytes(0));

                newProperties.Add(namedProperty1.Tag, namedProperty1);
                newProperties.Add(namedProperty2.Tag, namedProperty2);
                newProperties.Add(property.Tag, property);

                var updated = 0;
                foreach (var _ in testFolder.EnumerateMessagesEntryId()) updated++;

                testFolder.ChangeMessages(testFolder.EnumerateMessagesEntryId(), newProperties);

                Console.WriteLine($"Updated {updated} message(s) with {newProperties.Count} properties.");
                Console.WriteLine($"Saved to {pstPath}");
            }
        }

        // Builds a named-property tag: the property type in the low word, and the
        // named-property index (offset by 0x8000) in the high word.
        private static long GenerateNamedPropertyTag(long index, MapiPropertyType dataType)
        {
            return ((long)dataType | (0x8000 | index) << 16) & 0x00000000FFFFFFFF;
        }
    }
}
