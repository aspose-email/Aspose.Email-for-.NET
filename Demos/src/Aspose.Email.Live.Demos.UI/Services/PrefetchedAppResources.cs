using Aspose.Email.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.Services
{
    public class PrefetchedAppResources : IAppResources
    {
		Dictionary<string, string> _appResources = new Dictionary<string, string?>();

        public FlexibleResources AllResources { get; }
        public Uri RequestUri { get; }
        public string AppName { get; }
        public string ProductAppName { get; }

		public PrefetchedAppResources(FlexibleResources allResources,
			Uri requestUri,
			string appName)
		{
            AllResources = allResources;
            RequestUri = requestUri;

			AppName = allResources.GetValueOrDefault($"{appName}APPName") ?? appName;

			ProductAppName = "email" + AppName;
		}
		
        private string FindResource(string resourceName)
        {
			return AllResources.GetValueOrDefault(AppName + resourceName);
		}

        public bool ContainsResource(string resourceName)
        {
			if (_appResources.TryGetValue(resourceName, out string value))
			{
				return value != null;
			}
			else
			{
				return GetResourceOrDefault(resourceName) != null;
			}
        }

		public string GetResourceOrDefault(string resourceName, string defaultValue = null)
        {
			if (!_appResources.TryGetValue(resourceName, out string value))
			{
				value = FindResource(resourceName);
				_appResources.TryAdd(resourceName, value);
			}

			return value ?? defaultValue;
		}
    }
}
