using Aspose.Email.Mapi;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SetAdditionalMAPIProperties
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            MapiMessage msg = MapiMessage.FromFile(dataDir + "message.msg");

            // ExStart:SetAdditionalMAPIProperties
            // PT_MV_FLOAT, PT_MV_R4, mv.float
            IList values = new ArrayList();
            values.Add((float)1);
            values.Add((float)2);
            msg.SetProperty(new MapiProperty(0x23901004, values));

            // PT_MV_DOUBLE, PT_MV_R8
            values = new ArrayList();
            values.Add((double)1);
            values.Add((double)2);
            msg.SetProperty(new MapiProperty(0x23901005, values));

            // PT_MV_CURRENCY, mv.fixed.14.4
            values = new ArrayList();
            values.Add((decimal)123.34);
            values.Add((decimal)289.45);
            msg.SetProperty(new MapiProperty(0x23901006, values));

            // PT_MV_APPTIME
            values = new ArrayList();
            values.Add(30456.34);
            values.Add(40655.45);
            msg.SetProperty(new MapiProperty(0x23901007, values));

            // PT_MV_I8, PT_MV_LONGLONG
            values = new ArrayList();
            values.Add((long)30456);
            values.Add((long)40655);
            msg.SetProperty(new MapiProperty(0x23901014, values));

            // PT_MV_CLSID, mv.uuid
            values = new ArrayList();
            values.Add(Guid.NewGuid());
            values.Add(Guid.NewGuid());
            msg.SetProperty(new MapiProperty(0x23901048, values));

            // PT_MV_SHORT, PT_MV_I2, mv.i2
            values = new ArrayList();
            values.Add((short)1);
            values.Add((short)2);
            msg.SetProperty(new MapiProperty(0x23901002, values));

            // PT_MV_SYSTIME
            values = new ArrayList();
            values.Add(DateTime.Now);
            values.Add(DateTime.Now);
            msg.SetProperty(new MapiProperty(0x23901040, values));

            // PT_MV_BOOLEAN
            values = new ArrayList();
            values.Add(true);
            values.Add(false);
            msg.SetProperty(new MapiProperty(0x2390100b, values));

            // PT_MV_BINARY
            values = new ArrayList();
            values.Add(Guid.NewGuid().ToByteArray());
            values.Add(new byte[]{1,2,4,5,6,7,5,4,3,5,6,7,8,6,4,3,4,5,6,7,8,6,5,4,3,7,8,9,0,2,3,4,});
            msg.SetProperty(new MapiProperty(0x23901102, values));

            // PT_NULL
            msg.SetProperty(new MapiProperty(0x67400001, new byte[1]));
            MapiMessage message = new MapiMessage("sender@test.com", "recipient@test.com", "subj", "Body of test msg");

            // PT_MV_LONG
            values = new ArrayList();
            values.Add((int)4);
            MapiProperty property = new MapiProperty(message.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_MV_LONG), values);
            message.NamedPropertyMapping.AddNamedPropertyMapping(property, 0x00008028, new Guid("00062004-0000-0000-C000-000000000046"));
            message.SetProperty(property);

            // OR you can set the custom property (with the custom name)
            message = new MapiMessage("sender@test.com", "recipient@test.com", "subj", "Body of test msg");
            values = new ArrayList();
            values.Add((int)4);
            property = new MapiProperty(message.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_MV_LONG), values);
            message.AddCustomProperty(property, "customProperty");

            //PT_FLOAT
            //Please note that you need explicit cast to float value for this to work
            float floatValue = (float)123.456;
            MapiMessage newMsg = new MapiMessage();
            long floatTag = newMsg.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_FLOAT);
            Guid guid = Guid.NewGuid();
            MapiProperty newMapiProperty = new MapiProperty(floatTag, BitConverter.GetBytes(floatValue));
            newMsg.NamedPropertyMapping.AddNamedPropertyMapping(newMapiProperty, 12, guid);
            newMsg.SetProperty(newMapiProperty);

            // ExEnd:SetAdditionalMAPIProperties
        }
    }
}
