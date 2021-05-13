using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Helpers;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
    public partial class EmailService
    {
		public IEnumerable<SearchMatch> Search(Stream input, string pattern)
		{
			using (var mail = MapiHelper.GetMapiMessageFromStream(input))
			{
				var text = System.Net.WebUtility.HtmlDecode(mail.Body);

				var matches = Regex.Matches(text, pattern);

				if (matches.Count > 0)
				{

					int nodeCount = 0;
					foreach (Match match in matches)
					{
						int newIndex;
						int newLength;

						ExpandToNearWords(text, match.Index, match.Length, 2, out newIndex, out newLength);

						yield return new SearchMatch(
							nodeCount++,
							text.Substring(newIndex, match.Index - newIndex),
							text.Substring(match.Index, match.Length),
							text.Substring(match.Index + match.Length, newIndex + newLength - match.Index - match.Length));

					}
				}
			}
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
				if (char.IsWhiteSpace(text[i]))
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
				if (char.IsWhiteSpace(text[i]))
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
