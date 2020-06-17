using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Aspose.Email.Live.Demos.UI.LibraryHelpers
{
    /// <summary>
    /// Auxiliary methods for working with MAPI.
    /// Written temporarily until a new version is released.
    /// 
    /// EMAILNET-39534 Implement method of reading custom properties
    /// </summary>
    public static class MapiHelper
    {
        /// <summary>
        /// Get Content Id from attachment
        /// </summary>
        /// <param name="attach"></param>
        public static string GetContentId(this MapiAttachment attach)
		{
			return attach.GetPropertyString(KnownPropertyList.AttachContentId.Tag);
		}

		/// <summary>
        /// Set content id to attachment
        /// </summary>
        /// <param name="attach"></param>
        /// <param name="id"></param>
		public static void SetContentId(this MapiAttachment attach, object id)
		{
			attach.SetProperty(KnownPropertyList.AttachContentId, id);
		}

		///<Summary>
		/// Get Mapi Message From File
		///</Summary>
		/// <param name="fileNamePath"></param>
		public static Aspose.Email.Mapi.MapiMessage GetMapiMessageFromFile(string fileNamePath)
        {
            var formatInfo = Email.Tools.FileFormatUtil.DetectFileFormat(fileNamePath);

			switch (formatInfo.FileFormatType)
			{
				case Email.FileFormatType.Msg:
					return MapiMessage.Load(fileNamePath, new Email.MsgLoadOptions());
				case Email.FileFormatType.Eml:
					return MapiMessage.Load(fileNamePath, new Email.EmlLoadOptions());
				case Email.FileFormatType.Mht:
					return MapiMessage.Load(fileNamePath, new Email.MhtmlLoadOptions());
				case Email.FileFormatType.Tnef:
					return MapiMessage.Load(fileNamePath, new Email.TnefLoadOptions());
				default:
					return null;
			}
        }

		public static Aspose.Email.Mapi.MapiMessage GetMailMessageFromStream(Stream input)
		{
			var formatInfo = Email.Tools.FileFormatUtil.DetectFileFormat(input);
			input.Position = 0;

			switch (formatInfo.FileFormatType)
			{
				case Email.FileFormatType.Msg:
					return MapiMessage.Load(input, new Email.MsgLoadOptions());
				case Email.FileFormatType.Eml:
					return MapiMessage.Load(input, new Email.EmlLoadOptions());
				case Email.FileFormatType.Mht:
					return MapiMessage.Load(input, new Email.MhtmlLoadOptions());
				case Email.FileFormatType.Tnef:
					return MapiMessage.Load(input, new Email.TnefLoadOptions());
				default:
					return null;
			}
		}

		///<Summary>
		/// Get MailMessage From File
		///</Summary>
		/// <param name="fileNamePath"></param>
		public static MailMessage GetMailMessageFromFile(string fileNamePath)
        {
            var formatInfo = Email.Tools.FileFormatUtil.DetectFileFormat(fileNamePath);

            switch (formatInfo.FileFormatType)
            {
                case Email.FileFormatType.Msg:
                    return MailMessage.Load(fileNamePath, new Email.MsgLoadOptions());
                case Email.FileFormatType.Eml:
                    return MailMessage.Load(fileNamePath, new Email.EmlLoadOptions());
                case Email.FileFormatType.Mht:
                    return MailMessage.Load(fileNamePath, new Email.MhtmlLoadOptions());
                case Email.FileFormatType.Tnef:
                    return MailMessage.Load(fileNamePath, new Email.TnefLoadOptions());
                default:
                    return null;
            }
        }

        ///<Summary>
        /// Add Custom Property
        ///</Summary>
        /// <param name="msg"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>

        public static MapiProperty AddCustomProperty(this MapiMessage msg, string name, DateTime value)
        {
            var newProperty = MapiProperty.CreateMapiPropertyFromDateTime(msg.NamedPropertyMapping.GetNextAvailablePropertyId(MapiPropertyType.PT_SYSTIME), value);
            msg.AddCustomProperty(newProperty, name);
            return newProperty;
        }
		///<Summary>
		/// Add Custom Property
		///</Summary>
	    /// <param name="msg"></param>
		/// <param name="name"></param>
		/// <param name="value"></param>
		public static MapiProperty AddCustomProperty(this MapiMessage msg, string name, bool value)
        {
            return msg.AddCustomProperty(name, MapiPropertyType.PT_BOOLEAN, BitConverter.GetBytes(value));
        }
		///<Summary>
		/// Add Custom Property
		///</Summary>
		/// <param name="msg"></param>
		/// <param name="name"></param>
		/// <param name="value"></param>
		public static MapiProperty AddCustomProperty(this MapiMessage msg, string name, double value)
        {
            return msg.AddCustomProperty(name, MapiPropertyType.PT_DOUBLE, BitConverter.GetBytes(value));
        }
		///<Summary>
		/// Add Custom Property
		///</Summary>
		/// <param name="msg"></param>
		/// <param name="name"></param>
		/// <param name="value"></param>
		public static MapiProperty AddCustomProperty(this MapiMessage msg, string name, long value)
        {
            return msg.AddCustomProperty(name, MapiPropertyType.PT_LONG, BitConverter.GetBytes(value));
        }
		///<Summary>
		/// Add Custom Property
		///</Summary>
		/// <param name="msg"></param>
		/// <param name="name"></param>
		/// <param name="value"></param>
		public static MapiProperty AddCustomProperty(this MapiMessage msg, string name, string value)
        {
            return msg.AddCustomProperty(name, MapiPropertyType.PT_UNICODE, Encoding.Unicode.GetBytes(value));
        }
		///<Summary>
		/// Add Custom Property
		///</Summary>
		/// <param name="msg"></param>
		/// <param name="name"></param>
		/// <param name="type"></param>
		/// <param name="valueBytes"></param>
		public static MapiProperty AddCustomProperty(this MapiMessage msg, string name, MapiPropertyType type, byte[] valueBytes)
        {
            var newProperty = new MapiProperty(msg.NamedPropertyMapping.GetNextAvailablePropertyId(type), 0, valueBytes);
            msg.AddCustomProperty(newProperty, name);
            return newProperty;
        }
		///<Summary>
		/// Get Custom Properties
		///</Summary>
		/// <param name="msg"></param>		
		public static Dictionary<long, NamedMapiProperty> GetCustomProperties(this MapiMessage msg)
        {
            var customProperties = new Dictionary<long, NamedMapiProperty>();

            // 0x00008540 is equivalent of MapiNamedPropertyId.PidLidPropertyDefinitionStream
            long tag = GetTagFromNamedProperty(0x00008540, msg); 
            if (tag > 0)
            {
                byte[] data = msg.TryGetPropertyData(tag);
                List<string> names = GetCustomPropertiesNames(data);

                foreach (var name in names)
                {
                    long tagCustomProp = GetTagFromNamedProperty(name, msg);
                    if (tagCustomProp > 0)
                    {
                        MapiProperty customProp = msg.Properties[tagCustomProp];
                        if (customProp != null)
                        {
                            customProperties[tagCustomProp] = new NamedMapiProperty(name, customProp);
                        }
                    }
                }
            }

            return customProperties;
        }
		///<Summary>
		/// Get Tag From Named Property
		///</Summary>
		/// <param name="propertyId"></param>
		/// <param name="msg"></param>	
		private static long GetTagFromNamedProperty(long propertyId, MapiMessage msg)
        {
            foreach (MapiNamedProperty namedProperty in msg.NamedProperties.Values)
            {
                PidLidPropertyDescriptor lidDescriptor = namedProperty.Descriptor as PidLidPropertyDescriptor;
                if (lidDescriptor != null && lidDescriptor.LongId == propertyId)
                {
                    return namedProperty.Tag;
                }
            }

            return -1;
        }
		///<Summary>
		/// Get Tag From Named Property
		///</Summary>
		/// <param name="name"></param>
		/// <param name="msg"></param>
		private static long GetTagFromNamedProperty(string name, MapiMessage msg)
        {
            foreach (MapiNamedProperty namedProperty in msg.NamedProperties.Values)
            {
                if (namedProperty.NameId != null && name.Equals(namedProperty.NameId, StringComparison.OrdinalIgnoreCase))
                {
                    return namedProperty.Tag;
                }
            }

            return -1;
        }
		///<Summary>
		/// Get Custom Properties Names
		///</Summary>
		/// <param name="data"></param>
		
		private static List<string> GetCustomPropertiesNames(byte[] data)
        {
            List<string> result = new List<string>();
            if (data != null)
            {
                BinaryReader reader = new BinaryReader(new MemoryStream(data), Encoding.Unicode);
                string name = string.Empty;
                short ver = reader.ReadInt16();
                int numberOfProp = reader.ReadInt32();

                while (result.Count < numberOfProp && reader.PeekChar() > -1)
                {
                    //https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee415116%28v%3doffice.14%29

                    if (ver == 259) //0x0103 PropDefV2
                    {
                        name = GetCustomPropertyNameV2(reader);
                    }
                    else //0x0102 PropDefV1
                    {
                        name = GetCustomPropertyNameV1(reader);
                    }

                    if (!string.IsNullOrEmpty(name))
                    {
                        result.Add(name);
                    }
                }
            }

            return result;
        }
		///<Summary>
		/// Get Custom Property Name V1
		///</Summary>
		/// <param name="reader"></param>
		private static string GetCustomPropertyNameV1(BinaryReader reader)
        {
            //https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee415114%28v%3doffice.14%29

            string name = string.Empty;
            int PDO_IS_CUSTOM = 0x00000001;
            bool isCustom = (PDO_IS_CUSTOM & reader.ReadInt32()) == PDO_IS_CUSTOM;//Flags: DWORD (4 bytes)
            reader.ReadInt16();//VT: WORD (2 bytes)
            reader.ReadInt32();//DispId: DWORD (4 bytes)
            int nameIdLength = reader.ReadInt16();//NmidNameLength: WORD(2 bytes)
            name = new string(reader.ReadChars(nameIdLength));
            ReadPackedAnsiString(reader);//NameANSI
            ReadPackedAnsiString(reader);//FormulaANSI
            ReadPackedAnsiString(reader);//ValidationRuleANSI
            ReadPackedAnsiString(reader);//ValidationTextANSI
            ReadPackedAnsiString(reader);//ErrorANSI
            return isCustom ? name : string.Empty;
        }
		///<Summary>
		/// Get Custom Property Name V2
		///</Summary>
		/// <param name="reader"></param>
		private static string GetCustomPropertyNameV2(BinaryReader reader)
        {
            //https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee415114%28v%3doffice.14%29

            string name = GetCustomPropertyNameV1(reader);
            reader.ReadInt32();//InternalType: DWORD (4 bytes)
            int size = 0;
            while ((size = reader.ReadInt32()) > 0)//SkipBlocks
            {
                reader.ReadBytes(size);
            }
            return name;
        }
		///<Summary>
		/// Read Packed Ansi String
		///</Summary>
		/// <param name="reader"></param>
		private static void ReadPackedAnsiString(BinaryReader reader)
        {
            //https://docs.microsoft.com/ru-ru/office/client-developer/outlook/mapi/packedansistring-stream-structure

            int length = reader.ReadByte();
            if (length == 255)
            {
                length = reader.ReadInt16();
            }
            else if (length == 0)
            {
                return;
            }

            reader.ReadBytes(length);
        }
		///<Summary>
		/// Named Mapi Property
		///</Summary>
		
		public class NamedMapiProperty
        {///<Summary>
		 /// Get or Set Name property
		 ///</Summary>
			public string Name { get; private set; }
			///<Summary>
			/// Property Property
			///</Summary>
			public MapiProperty Property { get; private set; }
			///<Summary>
			/// Named Mapi Property
			///</Summary>
			public NamedMapiProperty(string name, MapiProperty customProp)
            {
                Name = name;
                Property = customProp;
            }
        }
    }
}
