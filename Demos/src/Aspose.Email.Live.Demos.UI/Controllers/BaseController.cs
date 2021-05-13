using Aspose.Email.Live.Demos.UI.Config;
using Aspose.Email.Live.Demos.UI.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
    public abstract class BaseController : Controller
	{
		public System.Uri RequestUrl => Request.GetUrl();
		
		public override void OnActionExecuted(ActionExecutedContext ctx)
		{
			base.OnActionExecuted(ctx);
			ViewBag.APIBasePath = Configuration.WebApiPath;
			ViewBag.Title = ViewBag.Title ?? Resources.GetValueOrDefault("ApplicationTitle");
		}

		private AppContext _atcContext;
		public AppContext AsposeToolsContext
		{
			get
			{
				if (_atcContext == null)
					_atcContext = new AppContext(Request.HttpContext);
				return _atcContext;
			}
		}

		FlexibleResources _resources;
		public FlexibleResources Resources
		{
			get
			{
				if (_resources == null)
					_resources = AsposeToolsContext.Resources;
				return _resources;
			}
		}
	}
}
