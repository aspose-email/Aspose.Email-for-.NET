using Aspose.Email.Live.Demos.UI.Models.Common;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class HeadersController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];
			   	
		public ActionResult Headers()
		{			
			var model = new ViewModel(this, "Headers")
			{
				ControlsView = "HeadersControls",
				MaximumUploadFiles = 10,
				DropOrUploadFileLabel = Resources["DropOrUploadFiles"]
				,
				ShowViewerButton = false
			};
			if (model.RedirectToMainApp)
				return Redirect("/email/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
