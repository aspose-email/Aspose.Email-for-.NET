using System;
using System.IO;
using System.Text;


namespace Aspose.Email.Live.Demos.UI.Services
{
    public static class ContentHandlingExtensionMethods
	{
		/// <summary>
		/// Create new output stream for resource by name, extension and part index.
		/// </summary>
		/// <param name="name">Name of resource</param>
		/// <param name="extension">Resource extension</param>
		/// <param name="partIndex">Resource part index (for partial content)</param>
		/// <returns>Empty stream for writing</returns>
		public static Stream CreateOutputStream(this IOutputHandler handler, string name, string extension, int? partIndex = null)
		{
			name = Path.ChangeExtension(name, null);

			if (extension == null)
				extension = "";

			if (!extension.IsNullOrEmpty() && !extension.StartsWith(".", StringComparison.Ordinal))
				extension = "." + extension;

			return handler.CreateOutputStream(handler.BuildName(name, extension, partIndex));
		}

		public static string RemoveUnsupportedCharacters(this string str)
		{
			var sb = new StringBuilder();

			foreach (char c in str)
			{
				if (c == '\\' ||
					c == '/' ||
					c == '"' ||
					c == '\'' ||
					c == '(' ||
					c == ')' ||
					c == '<' ||
					c == '>' ||
					c == '#' ||
					c == '$' ||
					c == '@' ||
					c == '%' ||
					c == '^' ||
					c == ' ')
					continue;

				sb.Append(c);
			}

			return sb.ToString();
		}
	}
}
