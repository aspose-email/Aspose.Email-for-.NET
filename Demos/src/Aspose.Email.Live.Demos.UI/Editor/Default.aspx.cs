using Aspose.Email.Live.Demos.UI.Config;
using System;
using System.Text;
using System.Web;
using System.Web.UI;
using Aspose.Email.Live.Demos.UI.Models;

namespace Aspose.Email.Live.Demos.UI.Editor
{
    public partial class Default : BasePage
    {
        public string FileName;
        public string FolderName;
        public string CallbackURL;
        public string DownloadOriginalURL;
        public string ProductName;
        public string AsposeEditorApp = Configuration.AsposeEmailLiveDemosPath + "api/EditorHelper/";
        public string FileDownloadLink = Configuration.AsposeEmailLiveDemosPath + "common/download";

        public string DownloadMainType;

        protected void Page_Load(object sender, EventArgs e)
        {
            Title = PageProductTitle + " " + Resources["EditorAPPName"];
			

            if (!IsPostBack)
            {
                ProductName = Resources["EditorAPPName"];
                Page.Title = Resources[Product + "EditorPageTitle"];
                Page.MetaDescription = Resources[Product + "EditorMetaDescription"];
                FileName = Request.QueryString["FileName"];
                FolderName = Request.QueryString["FolderName"];

                if (Request.QueryString["callbackURL"] != null)
                    CallbackURL = Request.QueryString["callbackURL"];
                else
                    CallbackURL = GetRouteUrl("AsposeEmailEditorRoute", new { Product });

                var url = new StringBuilder(Configuration.AsposeEmailLiveDemosPath + "common/download");
                url.Append("?FileName=");
                url.Append(HttpUtility.UrlEncode(FileName));
                url.Append("&FolderName=");
                url.Append(FolderName);
                DownloadOriginalURL = url.ToString();
            }

            var downloadItemsBuilder = new StringBuilder();

			downloadItemsBuilder.AppendLine(CreateListItem("MSG"));
			downloadItemsBuilder.AppendLine(CreateListItem("EML"));

			downloadItemsBuilder.AppendLine(CreateListItem("HTML"));

            litToDropdownItem.Text = downloadItemsBuilder.ToString();
        }

        string CreateListItem(string extension)
        {
            return $@"<li><a href=""#"" ng-click=""Download('.{extension.ToLowerInvariant()}')"">as {extension.ToUpperInvariant()}</a></li>";
        }
    }
}
