// Demonstrates how to load a MapiMessage and cast it to a MapiDistributionList.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadingDistributionListFromPst
    {
        public static void Run()
        {
            MapiMessage message = MapiMessage.Load(Data.Mapi/"NewGroup.msg");
            MapiDistributionList dlist = (MapiDistributionList)message.ToMapiMessageItem();
            Console.WriteLine("Members: " + dlist.Members.Count);
        }
    }
}
