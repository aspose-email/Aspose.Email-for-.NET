// Demonstrates how to set multi-valued MAPI properties of every supported type,
// and how to add named and custom properties to a message.

using System;
using System.Collections.Generic;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetAdditionalMAPIProperties
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");

            // Multi-valued properties take a list of values. The low word of the tag is
            // the property type - 0x1004 is PT_MV_FLOAT, 0x1005 PT_MV_DOUBLE, and so on.
            SetMultiValued(msg, 0x23901004, (float)1, (float)2);                       // PT_MV_FLOAT
            SetMultiValued(msg, 0x23901005, (double)1, (double)2);                     // PT_MV_DOUBLE
            SetMultiValued(msg, 0x23901006, (decimal)123.34, (decimal)289.45);         // PT_MV_CURRENCY
            SetMultiValued(msg, 0x23901007, 30456.34, 40655.45);                       // PT_MV_APPTIME
            SetMultiValued(msg, 0x23901014, (long)30456, (long)40655);                 // PT_MV_I8
            SetMultiValued(msg, 0x23901048, Guid.NewGuid(), Guid.NewGuid());           // PT_MV_CLSID
            SetMultiValued(msg, 0x23901002, (short)1, (short)2);                       // PT_MV_SHORT
            SetMultiValued(msg, 0x23901040, DateTime.Now, DateTime.Now);               // PT_MV_SYSTIME
            SetMultiValued(msg, 0x2390100b, true, false);                              // PT_MV_BOOLEAN
            SetMultiValued(msg, 0x23901102, Guid.NewGuid().ToByteArray(), new byte[] { 1, 2, 4, 5 }); // PT_MV_BINARY

            // PT_NULL carries no value of its own.
            msg.SetProperty(new MapiProperty(0x67400001, new byte[1]));

            var withNamedProperty = CreateWithNamedProperty();
            var withCustomProperty = CreateWithCustomProperty();
            var withFloatProperty = CreateWithFloatProperty();

            var outputPath = Data.Out/"SetAdditionalMAPIProperties_out.msg";
            msg.Save(outputPath);
            withNamedProperty.Save(Data.Out/"SetAdditionalMAPIProperties_named_out.msg");
            withCustomProperty.Save(Data.Out/"SetAdditionalMAPIProperties_custom_out.msg");
            withFloatProperty.Save(Data.Out/"SetAdditionalMAPIProperties_float_out.msg");

            Console.WriteLine($"Message now carries {msg.Properties.Count} properties.");
            Console.WriteLine($"Saved to {outputPath}");
            Console.WriteLine("Named, custom and float property samples saved alongside it.");
        }

        private static void SetMultiValued(MapiMessage msg, long tag, params object[] values)
        {
            msg.SetProperty(new MapiProperty(tag, new List<object>(values)));
        }

        // A named property is identified by an id within a property set GUID - this one
        // is PidLidTaskAssigner from the Outlook task property set.
        private static MapiMessage CreateWithNamedProperty()
        {
            var message = new MapiMessage("sender@test.com", "recipient@test.com", "subj", "Body of test msg");

            var property = new MapiProperty(
                message.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_MV_LONG),
                new List<object> { 4 });

            message.NamedPropertyMapping.AddNamedPropertyMapping(
                property, 0x00008028, new Guid("00062004-0000-0000-C000-000000000046"));
            message.SetProperty(property);

            return message;
        }

        // A custom property is identified by a name of your own instead of an id.
        private static MapiMessage CreateWithCustomProperty()
        {
            var message = new MapiMessage("sender@test.com", "recipient@test.com", "subj", "Body of test msg");

            var property = new MapiProperty(
                message.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_MV_LONG),
                new List<object> { 4 });

            message.AddCustomProperty(property, "customProperty");
            return message;
        }

        // PT_FLOAT needs an explicit cast to float, otherwise the literal is a double
        // and the resulting byte array would be the wrong length.
        private static MapiMessage CreateWithFloatProperty()
        {
            var message = new MapiMessage();

            var floatTag = message.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_FLOAT);
            var property = new MapiProperty(floatTag, BitConverter.GetBytes((float)123.456));

            message.NamedPropertyMapping.AddNamedPropertyMapping(property, 12, Guid.NewGuid());
            message.SetProperty(property);

            return message;
        }
    }
}
