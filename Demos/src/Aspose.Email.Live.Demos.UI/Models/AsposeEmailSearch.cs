using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Models.Common;

namespace Aspose.Email.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeEmailSearch class to perfrom search in email file
	///</Summary>
	public class AsposeEmailSearch : EmailBase
	{
		// Yellow Color
		private static readonly Color HiglightedColor = Color.FromArgb(246, 247, 146);
		///<Summary>
		/// IsValidRegex
		///</Summary>
		public static bool IsValidRegex(string pattern)
		{
			if (string.IsNullOrEmpty(pattern))
				return false;
			try
			{
				Regex.Match("", pattern);
			}
			catch (ArgumentException)
			{
				return false;
			}
			return true;
		}

		/// <summary>
		/// Redact documents
		/// </summary>
		
		public Response Search(InputFiles files, string query)
		{
			var statusValue = "OK";
			var statusCodeValue = 200;
			var fileProcessingErrorCode = FileProcessingErrorCode.OK;

			if (IsValidRegex(query))
			{			

				var processor = new CustomSingleOrZipFileProcessor()
				{
					
					CustomFilesProcessMethod= (string[] inputFilePaths, string outputFolderPath) =>
					{
						if (inputFilePaths.Length < 1)
							throw new Exception("No files provided");

						if (inputFilePaths.Length > 10)
							throw new Exception("Maximum 10 files allowed");

						var mailPairs = inputFilePaths.ToDictionary(x =>
						{
							var mail = MapiHelper.GetMapiMessageFromFile(x);

							if (mail == null)
								throw new Exception("Invalid file format " + x);

							return mail;
						}, x => x);

						var matchesCollections = mailPairs.ToDictionary(x => x, x => Regex.Matches(x.Key.Body, query));
						var doc = new Document();
						var builder = new DocumentBuilder(doc);
						builder.MoveToSection(0);

						if (matchesCollections.All(x => x.Value.Count == 0))
							throw new Exception("No search results");

						foreach (var pair in matchesCollections)
						{
							var mail = pair.Key.Key;
							var path = pair.Key.Value;
							var matches = pair.Value;

							if (matches.Count > 0)
							{
								var text = mail.Body;

								builder.Writeln("File: " + Path.GetFileName(path));
								builder.Writeln("Matches found: " + matches.Count);
								builder.Writeln();

								int nodeCount = 0;
								foreach (Match match in matches)
								{
									int newIndex;
									int newLength;

									ExpandToNearWords(text, match.Index, match.Length, 2, out newIndex, out newLength);

									var startText = text.Substring(newIndex, match.Index - newIndex);
									var endText = text.Substring(match.Index + match.Length, newIndex + newLength - match.Index - match.Length);

									var start = new Run(doc, startText.Replace("\n", " "));
									var value = new Run(doc, text.Substring(match.Index, match.Length));
									var end = new Run(doc, endText.Replace("\n", " "));

									value.Font.HighlightColor = HiglightedColor;

									builder.Writeln(++nodeCount + ":");

									builder.InsertNode(start);
									builder.InsertNode(value);
									builder.InsertNode(end);

									builder.Writeln();
									builder.Writeln();
								}
							}
						}

						doc.Save(Path.Combine(outputFolderPath, "Search results.docx"));

						fileProcessingErrorCode = FileProcessingErrorCode.OK;
					}
				};

				var responce = processor.Process(files.FolderName, files.FileName);

				if (responce.StatusCode == 200)
					return responce;

				fileProcessingErrorCode = FileProcessingErrorCode.NoSearchResults;
			}
			else
				fileProcessingErrorCode = FileProcessingErrorCode.WrongRegExp;

			return new Response
			{
				Status = statusValue,
				StatusCode = statusCodeValue,
				FileProcessingErrorCode = fileProcessingErrorCode
			};
		}

		

		private void ExpandToNearWords(string text, int index, int length, int wordsCount, out int newIndex, out int newLength)
		{
			int words = -1;
			// go left

			int i;
			int counter = 0;

			for (i = index; i > 0; i--)
			{
				counter++;
				if (Char.IsWhiteSpace(text[i]))
				{
					words++;
					counter = 0;
				}

				if (words == wordsCount || counter == 16)
					break;
			}

			newIndex = i;
			words = -1;
			counter = 0;

			// go right
			for (i = index + length; i < text.Length; i++)
			{
				counter++;
				if (Char.IsWhiteSpace(text[i]))
				{
					words++;
					counter = 0;
				}

				if (words == wordsCount || counter == 16)
					break;
			}

			newLength = i - newIndex;
		}
	}
}
