// Demonstrates how to enumerate named MAPI properties from a MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadNamedMapiProperties
    {
        public static void Run()
        {
            MapiMessage message = MapiMessage.Load(Data.Mapi/"message.msg");
            MapiPropertyCollection properties = message.NamedProperties;
            foreach (MapiNamedProperty mapiNamedProp in properties.Values)
            {
                switch (mapiNamedProp.NameId)
                {
                    case "TEST":
                    case "MYPROP":
                        Console.WriteLine("{0} = {1}", mapiNamedProp.NameId, mapiNamedProp.GetString());
                        break;
                }
            }
        }
    }
}
