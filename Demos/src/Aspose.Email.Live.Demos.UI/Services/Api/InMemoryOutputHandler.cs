using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services
{
	public class InMemoryOutputHandler : IOutputHandler
	{
		readonly Dictionary<string, MemoryStream> _resources = new Dictionary<string, MemoryStream>();
		readonly ReadOnlyDictionary<string, MemoryStream> _readOnlyResources;
		public IReadOnlyDictionary<string, MemoryStream> Resources => _readOnlyResources;

		public InMemoryOutputHandler()
		{
			_readOnlyResources = new ReadOnlyDictionary<string, MemoryStream>(_resources);
		}

		public Stream CreateOutputStream(string name)
		{
			return new StreamProxy(_resources[name] = new MemoryStream());
		}

		public Stream OpenReadStream(string name)
		{
			if (_resources.TryGetValue(name, out MemoryStream stream))
			{
				stream.Position = 0;
				return new StreamProxy(stream);
			}

			throw new ArgumentException(nameof(name));
		}

		public void RemoveResource(string name)
		{
			_resources.Remove(name);
		}

		public string BuildName(string nameWithoutExtension, string extension, int? partIndex = null)
		{
			return nameWithoutExtension + (partIndex?.ToString() ?? "") + extension;
		}
	}
}
