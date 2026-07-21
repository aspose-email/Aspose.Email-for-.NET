// Demonstrates how to remove the password property from a PST file, so that Outlook
// stops prompting for it.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RemovingPaswordProperty
    {
        public static void Run()
        {
            // Work on a copy: examples must never modify the shared input data.
            var pstPath = Data.Out/"RemovingPaswordProperty_out.pst";
            File.Copy(Data.Mapi/"PersonalStorage1.pst", pstPath, true);

            using (var personalStorage = PersonalStorage.FromFile(pstPath))
            {
                if (personalStorage.Store.Properties.ContainsKey(MapiPropertyTag.PR_PST_PASSWORD))
                {
                    // Setting the property to zero clears the password.
                    var property = new MapiProperty(MapiPropertyTag.PR_PST_PASSWORD, BitConverter.GetBytes((long)0));
                    personalStorage.Store.SetProperty(property);

                    Console.WriteLine($"Password property removed. Saved to {pstPath}");
                }
                else
                {
                    Console.WriteLine("This PST file has no password property set.");
                }
            }
        }
    }
}
