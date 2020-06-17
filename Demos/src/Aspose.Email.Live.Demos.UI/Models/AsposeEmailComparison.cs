using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using static Aspose.Email.Live.Demos.UI.Helpers.LongestCommonSubsequenceUtils;
using Aspose.Email.Live.Demos.UI.Models.Common;

namespace Aspose.Email.Live.Demos.UI.Models
{
    public class AsposeEmailComparison : EmailBase
	{
		///<Summary>
		/// Merge method to merge word document
		///</Summary>
		
		public Response Compare(InputFiles inputFiles)
		{
		    (string folder, string[] files) = (null, null);
			folder = inputFiles.FolderName;
			files = inputFiles.FileName;
			try
			{				
				return  Compare(Path.GetFileName(files[0]), Path.GetFileName(files[1]), folder);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}

		///<Summary>
		/// Compare method to compare emails 
		///</Summary>
		public Response Compare(string fileName1, string fileName2, string folderName)
        {
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();

            var comparedDocument = string.Format("{0} Compare To {1}.msg",
                Path.GetFileNameWithoutExtension(fileName1), Path.GetFileNameWithoutExtension(fileName2));

            return  Process(this.GetType().Name, comparedDocument, folderName, ".msg", false, false,
				 "Compare",
                (inFilePath, outPath, zipOutFolder) =>
                {
                    var email1 = MapiHelper.GetMapiMessageFromFile(Path.Combine(Config.Configuration.WorkingDirectory, folderName, fileName1));
                    var email2 = MapiHelper.GetMapiMessageFromFile(Path.Combine(Config.Configuration.WorkingDirectory, folderName, fileName2));

                    var body1 = email1.Body.Split(new string[] { " " }, StringSplitOptions.None).SelectMany(SplitOfNewLine).ToArray();
                    var body2 = email2.Body.Split(new string[] { " " }, StringSplitOptions.None).SelectMany(SplitOfNewLine).ToArray();

                    var lcsMatrix = GetMatrix(body1, body2, StringComparer.Ordinal);

                    var list = new List<LSTnode<string>>();

                    ListDiffTreeFromBacktrackMatrix(list, lcsMatrix, body1, body2, StringComparer.Ordinal);

                    email2.SetBodyContent(BuildBodyDiff(list), Email.Mapi.BodyContentType.Html);

                    // Alternative comparison. Better visuals but may broke mail render in Outlook.
                    //var htmlDiff = new Helpers.Html.HtmlDiff(email1.BodyHtml, email2.BodyHtml);
                    //email2.SetBodyContent(htmlDiff.Build(), Email.Mapi.BodyContentType.Html);

                    email2.Save(outPath);
                });
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
