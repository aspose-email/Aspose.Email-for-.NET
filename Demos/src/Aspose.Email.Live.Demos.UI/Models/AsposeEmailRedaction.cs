using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Parsing.Html;
using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Models.Common;

namespace Aspose.Email.Live.Demos.UI.Models
{
	public class AsposeEmailRedaction : EmailBase
	{
        private List<string> _allowedTypes = new List<string> { /*"Boolean", "DateTime", "Double", "Number", "Int32",*/ "String" };
        bool IsAllowedType(string propertyType)
        {
            return _allowedTypes.Contains(propertyType);
        }

		
		public Response Redact(InputFiles inputFiles, string outputType, string searchQuery, string replaceText, bool caseSensitive, bool text, bool comments, bool metadata)
		{
			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFilesProcessMethod = (string[] inputFilePaths, string outputFolderPath) =>
				{
					if (inputFilePaths.Length < 1)
						throw new Exception("No files provided");

					if (inputFilePaths.Length > 10)
						throw new Exception("Maximum 10 files allowed");

					var mails = inputFilePaths.ToDictionary(x =>
					{
						var mail = MapiHelper.GetMapiMessageFromFile(x);

						if (mail == null)
							throw new Exception("Invalid file format");

						return mail;
					}, x => x).ToArray();

					var regex = new Regex(searchQuery, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);

					foreach (var mailAndPath in mails)
					{
						var mail = mailAndPath.Key;
						var path = mailAndPath.Value;

						if (metadata)
						{
							Type t = mail.GetType();

							foreach (var prop in t.GetProperties())
							{
								if (prop.CanWrite && IsAllowedType(prop.PropertyType.Name))
								{
									var valueType = prop.PropertyType.Name;

									if (prop.Name == nameof(MapiMessage.Body) || prop.Name == nameof(MapiMessage.BodyHtml) || prop.Name == nameof(MapiMessage.BodyRtf))
										continue;

									if (valueType == "String")
									{
										var value = (string)prop.GetValue(mail);

										if (value == null)
											continue;

										var matches = regex.Matches(value);

										int offset = 0;

										for (int m = 0; m < matches.Count; m++)
										{
											value = value.Remove(matches[m].Index + offset, matches[m].Length).Insert(matches[m].Index + offset, replaceText);
											offset = offset + replaceText.Length - matches[m].Length;
										}

										prop.SetValue(mail, value);
									}

								}
							}

							foreach (var pair in MapiHelper.GetCustomProperties(mail))
							{
								var name = pair.Value.Name;
								var prop = pair.Value.Property;

								var valueType = ConvertMailPropertyType((MapiPropertyType)prop.DataType);

								if (valueType == "String")
								{
									var value = (string)prop.GetValue();

									if (value == null)
										continue;

									var matches = regex.Matches(value);

									int offset = 0;

									for (int m = 0; m < matches.Count; m++)
									{
										value = value.Remove(matches[m].Index + offset, matches[m].Length).Insert(matches[m].Index + offset, replaceText);
										offset = offset + replaceText.Length - matches[m].Length;
									}

									mail.SetProperty(prop.Descriptor, value);
								}
							}
						}

						mail.SetBodyContent(TraverseHtml(mail.BodyHtml, regex, replaceText, comments, metadata), BodyContentType.Html);
						mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(path) + " Redacted.msg"), MsgSaveOptions.DefaultMsgUnicode);
					}
				}
			};

			(string folder, string[] files) = (null, null);
			folder = inputFiles.FolderName;
			files = inputFiles.FileName;

			try
			{
				
				return processor.Process(folder, files);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}

		public Response Redact(string fileName, string folderName, string searchQuery, string replaceText, bool caseSensitive, bool text, bool comments, bool metadata)
        {
            var processor = new CustomSingleOrZipFileProcessor()
            {
                
                CustomFileProcessMethod = (string inputFilePath, string outputFolderPath) =>
                {
                    var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

                    if (mail == null)
                        throw new Exception("Invalid file format");

                    var regex = new Regex(searchQuery, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);

                    if (metadata)
                    {
                        Type t = mail.GetType();

                        foreach (var prop in t.GetProperties())
                        {
                            if (prop.CanWrite && IsAllowedType(prop.PropertyType.Name))
                            {
                                var valueType = prop.PropertyType.Name;

                                if (prop.Name == nameof(MapiMessage.Body) || prop.Name == nameof(MapiMessage.BodyHtml) || prop.Name == nameof(MapiMessage.BodyRtf))
                                    continue;

                                if (valueType == "String")
                                {
                                    var value = (string)prop.GetValue(mail);

									if (value == null)
                                        continue;

                                    var matches = regex.Matches(value);

                                    int offset = 0;

                                    for (int m = 0; m < matches.Count; m++)
                                    {
                                        value = value.Remove(matches[m].Index + offset, matches[m].Length).Insert(matches[m].Index + offset, replaceText);
                                        offset = offset + replaceText.Length - matches[m].Length;
                                    }

                                    prop.SetValue(mail, value);
                                }

                            }
                        }
						
                        foreach (var pair in MapiHelper.GetCustomProperties(mail))
                        {
                            var name = pair.Value.Name;
                            var prop = pair.Value.Property;

                            var valueType = ConvertMailPropertyType((MapiPropertyType)prop.DataType);

                            if (valueType == "String")
                            {
                                var value = (string)prop.GetValue();

                                if (value == null)
                                    continue;

                                var matches = regex.Matches(value);

                                int offset = 0;

                                for (int m = 0; m < matches.Count; m++)
                                {
                                    value = value.Remove(matches[m].Index + offset, matches[m].Length).Insert(matches[m].Index + offset, replaceText);
                                    offset = offset + replaceText.Length - matches[m].Length;
                                }

                                mail.SetProperty(prop.Descriptor, value);
                            }
                        }
                    }

                    mail.SetBodyContent(TraverseHtml(mail.BodyHtml, regex, replaceText, comments, metadata), BodyContentType.Html);
					mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + " Redacted.msg"), MsgSaveOptions.DefaultMsgUnicode);
                }
            };

            return processor.Process(folderName, fileName);
        }

        string ConvertMailPropertyType(MapiPropertyType type)
        {
            switch (type)
            {
                case MapiPropertyType.PT_BOOLEAN:
                    return "Boolean";
                case MapiPropertyType.PT_DOUBLE:
                    return "Double";
                case MapiPropertyType.PT_LONG:
                    return "Number";
                case MapiPropertyType.PT_SYSTIME:
                    return "DateTime";
                case MapiPropertyType.PT_UNICODE:
                    return "String";
                default:
                    return type.ToString();
            }
        }

        string TraverseHtml(string bodyHtml, Regex regex, string replace, bool comments, bool metadata)
        {
            var parser = new HtmlLexemmeParser();

            var lexemmes = parser.ParseString(bodyHtml);

            var redactedHtmlBuilder = new StringBuilder();

            for (int i = 0; i < lexemmes.Count; i++)
            {
                var lex = lexemmes[i];

                if (lex.Type != nameof(HtmlLexemmeType.Text) && (!comments || lex.Type != nameof(HtmlLexemmeType.Comment)))
                {
                    redactedHtmlBuilder.Append(lex.StringView);
                    continue;
                }

                var matches = regex.Matches(lex.StringView);
                var str = lex.StringView;

                if (lex.Type == nameof(HtmlLexemmeType.Comment))
                    str = str.Substring(3, str.Length - 5);

                int offset = 0;

                for (int m = 0; m < matches.Count; m++)
                {
                    str = str.Remove(matches[m].Index + offset, matches[m].Length).Insert(matches[m].Index + offset, replace);
                    offset = offset + replace.Length - matches[m].Length;
                }

                redactedHtmlBuilder.Append(str);

            }

            return redactedHtmlBuilder.ToString();
        }
    }
}
