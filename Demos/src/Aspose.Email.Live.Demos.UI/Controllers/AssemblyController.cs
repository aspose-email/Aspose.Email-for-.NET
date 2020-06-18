using Aspose.Email.Live.Demos.UI.Models.Common;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;
using System.Net.Http;

namespace Aspose.Email.Live.Demos.UI.Controllers
{
	public class AssemblyController : BaseController
	{
		public override string Product => (string)RouteData.Values["product"];		


			public ActionResult Assembly()
		{
			var model = new ViewModel(this, "Assembly")
			{				
				MaximumUploadFiles = 1,
				ControlsView = "AssemblyControls",
				DefaultFileBlockDisabled = true,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"],
				UploadAndRedirect = true
			};
			if (model.RedirectToMainApp)
				return Redirect("/email/" + model.AppName.ToLower());
			return View(model);		
			
		}	

	}
}
