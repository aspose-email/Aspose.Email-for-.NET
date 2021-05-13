using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Services;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Caching.Memory;

namespace Aspose.Email.Live.Demos.UI
{
    public abstract class BaseRootPageModel : PageModel
    {
        public string RequestUrlAbsoluteUri => Request.GetDisplayUrl();

        private AppContext _appContext;

        public AppContext AppContext
        {
            get
            {
                if (_appContext == null) 
                    _appContext = new AppContext(Request.HttpContext);
                return _appContext;
            }
        }

        public string Locale
        {
            get
            {
                var routeValues = RouteData.Values;
                if (routeValues.ContainsKey("locale") && !string.IsNullOrEmpty(routeValues["locale"].ToString()))
                {
                    return routeValues["locale"].ToString().ToUpper();
                }
                else
                {
                    return "EN";
                }
            }
        }

        public abstract string AppName { get; }
        public IAppResources AppResources { get; private set; }

        public virtual void OnGet()
        {
            // ---- Prepare SEO resources
            var locale = Locale;
            var appName = AppName;
            var uri = Request.GetUrl();
            var resources = AppContext.Resources;

            var appResourcesKey = $"appresources_{appName}___{locale}";

            var appResources = HttpContext.GetCache().GetOrCreate<IAppResources>(appResourcesKey, entry =>
            {
                entry.Priority = CacheItemPriority.NeverRemove;
                return new PrefetchedAppResources(resources, uri, appName);
            });

            AppResources = new PrefetchedAppResources(resources, uri, appName);
        }

        public string this[string key] => AppResources.GetResourceOrDefault(key);
    }
}
