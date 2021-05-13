using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Aspose.Email.Live.Demos.UI.Services
{
    public class ViewerImagesOutputHandler : IOutputHandler
	{
		readonly string _stringFormat;

		readonly Dictionary<string, int> _nameToPositionMap = new Dictionary<string, int>();
		readonly Dictionary<string, MemoryStream> _nameToStreamMap = new Dictionary<string, MemoryStream>();


		public string[] GetOrderedPageNames() => _nameToPositionMap.OrderBy(x => x.Value).Select(x => x.Key).ToArray();

		public ViewerImagesOutputHandler(string format)
		{
			_stringFormat = format;
		}

		public int ImagesCount { get; private set; }

		public Stream CreateOutputStream(string name)
		{
			PrepareName(ref name);

            var stream = new MemoryStream();
			_nameToStreamMap.Add(name, stream);
			return new StreamProxy(stream);
		}

		public string GetPathForStream(string name)
		{
			throw new NotSupportedException();
		}

		public Stream OpenReadStream(string name)
		{
			PrepareName(ref name);
			var stream = _nameToStreamMap[name];
			stream.Position = 0;
			return new StreamProxy(stream);
		}

		public void RemoveResource(string name)
		{
			PrepareName(ref name);
			_nameToStreamMap.Remove(name);
		}

		private void PrepareName(ref string name)
		{
			if (name.EndsWith(".jpg", System.StringComparison.OrdinalIgnoreCase) || name.EndsWith(".jpeg", System.StringComparison.OrdinalIgnoreCase))
			{
				name = name.RemoveUnsupportedCharacters();

				if (_nameToPositionMap.TryGetValue(name, out int value))
				{
					name = string.Format(_stringFormat, value, name);
				}
				else
				{
					name = string.Format(_stringFormat, _nameToPositionMap[name] = ImagesCount++, name);
				}
			}
		}

		public string BuildName(string nameWithoutExtension, string extension, int? partIndex = null)
		{
			if (partIndex == null)
				return nameWithoutExtension + extension;

			return nameWithoutExtension + extension + "-part"+ partIndex.ToString();
		}
	}
}

