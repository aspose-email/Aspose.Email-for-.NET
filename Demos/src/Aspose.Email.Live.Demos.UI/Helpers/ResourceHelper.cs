using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    public static class ResourceHelper
    {
        static Regex FormatsRegex = new Regex(@"(^|,+|\s+|\.+|/|\\|\|)(msg|eml|pst|ost|mbox|pdf|mhtml|jpeg|jpg|png|bmp|tiff|pptx|ppt|rtf|docx|docm|dotx|dotm|odt|ott|txt|xps|pcl|ps|html|doc)(,+|\s+|\.+|/|\\|\||$)", RegexOptions.IgnoreCase);

		/// <summary>
		/// Inject links to format description. Formats list:
		/// EML, MSG, PST, OST, MBOX, PDF, HTML, JPG, PNG, BMP, TIFF, DOC, PPT, RTF, DOCX, DOCM, DOTX, DOTM, ODT, OTT, TXT, XPS, PCL, PS, MHTML.
		/// </summary>
		public static string InjectFormatLinks(string value)
		{
			var matches = FormatsRegex.Matches(value);

			if (matches.Count > 0)
			{
				var builder = new StringBuilder(value);

				for (int i = 0; i < matches.Count; i++)
				{
					var match = matches[matches.Count - 1 - i];

					var text = match.Value.Trim(' ', ',', '/', '\\', '.', '|');

					if (text.IsNullOrWhiteSpace())
						continue;

					string link = null;

					switch (text.ToLowerInvariant())
					{
						case "eml": link = @"https://docs.fileformat.com/email/eml/"; break;
						case "msg": link = @"https://docs.fileformat.com/email/msg/"; break;
						case "pst": link = @"https://docs.fileformat.com/email/pst/"; break;
						case "ost": link = @"https://docs.fileformat.com/email/ost/"; break;
						case "mbox": link = @"https://docs.fileformat.com/email/mbox/"; break;
						case "pdf": link = @"https://docs.fileformat.com/pdf/"; break;
						case "html": link = @"https://docs.fileformat.com/web/html/"; break;
						case "jpeg": link = @"https://docs.fileformat.com/image/jpeg/"; break;
						case "jpg": link = @"https://docs.fileformat.com/image/jpeg/"; break;
						case "png": link = @"https://docs.fileformat.com/image/png/"; break;
						case "bmp": link = @"https://docs.fileformat.com/image/bmp/"; break;
						case "tiff": link = @"https://docs.fileformat.com/image/tiff/"; break;
						case "doc": link = @"https://docs.fileformat.com/word-processing/doc/"; break;
						case "ppt": link = @"https://docs.fileformat.com/presentation/ppt/"; break;
						case "pptx": link = @"https://docs.fileformat.com/presentation/pptx/"; break;
						case "rtf": link = @"https://docs.fileformat.com/word-processing/rtf/"; break;
						case "docx": link = @"https://docs.fileformat.com/word-processing/docx/"; break;
						case "docm": link = @"https://docs.fileformat.com/word-processing/docm/"; break;
						case "dotx": link = @"https://docs.fileformat.com/word-processing/dotx/"; break;
						case "dotm": link = @"https://docs.fileformat.com/word-processing/dotm/"; break;
						case "odt": link = @"https://docs.fileformat.com/word-processing/odt/"; break;
						case "ott": link = @"https://docs.fileformat.com/word-processing/ott/"; break;
						case "txt": link = @"https://docs.fileformat.com/word-processing/txt/"; break;
						case "xps": link = @"https://docs.fileformat.com/page-description-language/xps/"; break;
						case "pcl": link = @"https://docs.fileformat.com/page-description-language/pcl/"; break;
						case "ps": link = @"https://docs.fileformat.com/page-description-language/ps/"; break;
						case "mhtml": link = @"https://docs.fileformat.com/web/mhtml/"; break;
						default:
							break;
					}

					if (link == null)
						continue;

					builder.Remove(match.Index, match.Length);
					builder.Insert(match.Index, match.Value.Replace(text, $@"<a target=""_blank"" style=""color: #655ed2!important"" rel=""nofollow"" href=""{link}"">{text}</a>"));
				}

				value = builder.ToString();
			}

			return value;
		}

		public static Dictionary<string, string> ReadAppResources(string language = null)
		{
			var defaultResourcesFilePath = Path.Combine(Startup.WebHostEnvironment.ContentRootPath, "App_Data/resources_EN" + ".xml");

			var resources = new Dictionary<string, string>();

			LoadDocumentToDictionary(defaultResourcesFilePath, resources);

			if (!language.IsNullOrEmpty())
			{
				var resourcesByLanguageFilePath = Path.Combine(Startup.WebHostEnvironment.ContentRootPath, "App_Data/resources_" + language + ".xml");

				if (File.Exists(resourcesByLanguageFilePath))
					LoadDocumentToDictionary(defaultResourcesFilePath, resources);
			}

			return resources;
		}

		private static void LoadDocumentToDictionary(string filePath, Dictionary<string, string> resources)
		{
			var document = new XmlDocument();

			document.Load(filePath);

			var nodes = document.SelectNodes("resources/res");

			foreach (XmlNode n in nodes)
				resources[n.Attributes["name"].Value] = n.InnerText;
		}
	}
}
