// Demonstrates how to check whether a PST file is password protected by inspecting its PR_PST_PASSWORD property.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CheckPasswordProtection
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"passwordprotectedPST.pst"))
            {
                Console.WriteLine($"PST is protected: {IsPasswordProtected(pst)}");
            }
        }

        private static bool IsPasswordProtected(PersonalStorage pst)
        {
            if (pst.Store.Properties.ContainsKey(MapiPropertyTag.PR_PST_PASSWORD))
                return pst.Store.Properties[MapiPropertyTag.PR_PST_PASSWORD].GetLong() != 0;

            return false;
        }
    }
}
