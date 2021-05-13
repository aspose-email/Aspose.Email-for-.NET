using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Models
{
    public class FlexibleResources
	{
		Dictionary<string, string> _resources;

		public FlexibleResources(Dictionary<string, string> resourcesFile)
		{
			_resources = resourcesFile;
		}

		public bool ContainsKey(string key)
		{
			return _resources.ContainsKey(key);
		}

		public string GetValueOrDefault(string key)
		{
			if (_resources.TryGetValue(key, out string value))
				return value;

			return null;
		}
	}
}
