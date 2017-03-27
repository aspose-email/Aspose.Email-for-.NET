using System;
using System.Text;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class UpdatePSTCustomProperites
    {
        // ExStart:UpdatePSTCustomProperites
        public static void Run()
        {
            // Load the Outlook file
            string dataDir = RunExamples.GetDataDir_Outlook() + "Outlook.pst";
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir))
            {
                FolderInfo testFolder = personalStorage.RootFolder.GetSubFolder("Inbox");

                // Create the collection of message properties for adding or updating
                MapiPropertyCollection newProperties = new MapiPropertyCollection();

                // Normal,  Custom and PidLidLogFlags named  property
                MapiProperty property = new MapiProperty(MapiPropertyTag.PR_ORG_EMAIL_ADDR_W,Encoding.Unicode.GetBytes("test_address@org.com"));
                MapiProperty namedProperty1 = new MapiNamedProperty(GenerateNamedPropertyTag(0, MapiPropertyType.PT_LONG),"ITEM_ID",Guid.NewGuid(),BitConverter.GetBytes(123));
                MapiProperty namedProperty2 = new MapiNamedProperty(GenerateNamedPropertyTag(1, MapiPropertyType.PT_LONG),0x0000870C,new Guid("0006200A-0000-0000-C000-000000000046"),BitConverter.GetBytes(0));
                newProperties.Add(namedProperty1.Tag, namedProperty1);
                newProperties.Add(namedProperty2.Tag, namedProperty2);
                newProperties.Add(property.Tag, property);
                testFolder.ChangeMessages(testFolder.EnumerateMessagesEntryId(), newProperties);
            }          
        }

        private static long GenerateNamedPropertyTag(long index, MapiPropertyType dataType)
        {
            return ((long)(long)dataType | (0x8000 | index) << 16) & 0x00000000FFFFFFFF;
        }
        // ExEnd:UpdatePSTCustomProperites
    }
}