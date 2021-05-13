using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static Aspose.Email.Live.Demos.UI.Helpers.LongestCommonSubsequenceUtils;

namespace Aspose.Email.Live.Demos.UI.Services.Email
{
	public partial class EmailService
    {
		public void Compare(Stream mail1Stream, Stream mail2Stream, string shortResourceNameWithExtension, IOutputHandler handler)
		{
			using (var email1 = MapiHelper.GetMapiMessageFromStream(mail1Stream))
			using (var email2 = MapiHelper.GetMapiMessageFromStream(mail2Stream))
			{
				var body1 = email1.Body.Split(new string[] { " " }, StringSplitOptions.None).SelectMany(SplitOfNewLine).ToArray();
				var body2 = email2.Body.Split(new string[] { " " }, StringSplitOptions.None).SelectMany(SplitOfNewLine).ToArray();

				var lcsMatrix = GetMatrix(body1, body2, StringComparer.Ordinal);

				var list = new List<LSTnode<string>>();

				ListDiffTreeFromBacktrackMatrix(list, lcsMatrix, body1, body2, StringComparer.Ordinal);

				email2.SetBodyContent(BuildBodyDiff(list), BodyContentType.Html);

				// Alternative comparison. Better visuals but may broke mail render in Outlook.
				//var htmlDiff = new Helpers.Html.HtmlDiff(email1.BodyHtml, email2.BodyHtml);
				//email2.SetBodyContent(htmlDiff.Build(), Email.Mapi.BodyContentType.Html);

				using (var output = handler.CreateOutputStream(Path.ChangeExtension(shortResourceNameWithExtension, ".Compared" + Path.GetExtension(shortResourceNameWithExtension))))
					email2.Save(output);
			}
		}

		IEnumerable<string> SplitOfNewLine(string value)
		{
			if (value.IndexOf(Environment.NewLine, StringComparison.Ordinal) == -1)
				return new[] { value };

			var split = value.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

			var newSplit = new string[split.Length + split.Length - 1];

			for (int i = 0; i < split.Length; i++)
			{
				newSplit[i << 1] = split[i];

				if (i < split.Length - 1)
					newSplit[(i << 1) + 1] = Environment.NewLine;
			}

			return newSplit;
		}

		string BuildBodyDiff(IList<LSTnode<string>> list)
		{
			var builder = new StringBuilder();

			builder.Append(@"<body><pre>");

			var currentMode = NodeMode.NotChanged;
			var currentSequence = new StringBuilder();

			for (int i = 0; i < list.Count; i++)
			{
				var node = list[i];

				var value = node.Value;

				if (node.Mode != currentMode)
				{
					AppendSequence(currentMode, builder, currentSequence.ToString());

					currentSequence.Clear();

					if (currentSequence.Length > 0)
						currentSequence.Append(" ");

					currentSequence.Append(value);
					currentMode = node.Mode;
				}
				else
				{
					if (currentSequence.Length > 0)
						currentSequence.Append(" ");

					currentSequence.Append(value);
				}
			}

			if (currentSequence.Length != 0)
				AppendSequence(currentMode, builder, currentSequence.ToString());

			builder.Append("</pre></body>");

			return builder.ToString();
		}

		void AppendSequence(NodeMode currentMode, StringBuilder builder, string sequence)
		{
			sequence = sequence
				.Replace("&", "&amp;")
				.Replace(">", "&gt;")
				.Replace("<", "&lt;")
				.Replace("\"", "&quot;")
				.Replace("/", "&#x2F;");

			switch (currentMode)
			{
				case NodeMode.NotChanged:
					break;
				case NodeMode.Added:
					sequence = $@"<font color=""blue""><b>{sequence}</b></font> ";
					break;
				case NodeMode.Removed:
					sequence = $@"<font color=""red""><s>{sequence}</s></font> ";
					break;
				default:
					break;
			}

			builder.Append(sequence);
			builder.Append(" ");
		}
	}
}
