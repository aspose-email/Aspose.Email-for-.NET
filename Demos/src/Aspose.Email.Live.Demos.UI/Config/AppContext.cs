using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Caching.Memory;

namespace Aspose.Email.Live.Demos.UI.Config
{
    public class AppContext
	{
        private const string DefaultResourcesCacheKey = "resources";

        HttpContext _context;
        FlexibleResources _resources;

        public AppContext(HttpContext context)
		{
            _context = context;
		}

        public IMemoryCache Cache => _context.GetCache();
        protected string Locale => _context.Request.GetUrl().Host.StartsWith("zh.", System.StringComparison.OrdinalIgnoreCase) ? "ZH" : "EN";


		public FlexibleResources Resources
		{
            get
            {
                if (_resources == null)
                {
                    var locale = Locale;
                    var key = $"{DefaultResourcesCacheKey}_{locale}";

                    if (!Cache.TryGetValue(key, out _resources))
                    {
                        var options = new MemoryCacheEntryOptions()
                        {
                            Priority = CacheItemPriority.NeverRemove,
                            AbsoluteExpiration = null,
                            SlidingExpiration = null
                        };

                        options.PostEvictionCallbacks.Add(new PostEvictionCallbackRegistration()
                        {
                            EvictionCallback = (key, value, reason, state) => InitFlexibleResources(locale)
                        });

                        _resources = Cache.Set(key, InitFlexibleResources(locale));
                    }
                }

                return _resources;
            }
		}

        private FlexibleResources InitFlexibleResources(string locale)
        {
            var baseResources = ResourceHelper.ReadAppResources(locale);
            return new FlexibleResources(baseResources);
        }
    }
}
