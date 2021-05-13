using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Api.Storage;
using Aspose.Email.Live.Demos.UI.Services.Email;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Net.Http.Formatting;

namespace Aspose.Email.Live.Demos.UI
{
    public class Startup
	{
		public static IConfiguration Configuration { get; private set; }
		public static IWebHostEnvironment WebHostEnvironment { get; private set; }

		public static ILogger<DefaultExceptionFilterAttribute> ExceptionFilterLogger { get; private set; }

		public Startup(IConfiguration configuration, IWebHostEnvironment env)
		{
			Configuration = configuration;
			Config.Configuration.SetConfiguration(configuration);
		}

		// This method gets called by the runtime. Use this method to add services to the container.
		public void ConfigureServices(IServiceCollection services)
		{
			services.AddSingleton(Configuration);

			services.AddCors();

			services.AddSingleton<IContentNegotiator, DefaultContentNegotiator>();

			services.AddHttpClient();
			services.AddSingleton<IStorageService, FileStorageService>();
			services.AddSingleton<IEmailService, EmailService>();

			services.Configure<KestrelServerOptions>(options =>
			{
				options.AllowSynchronousIO = true;
			});

			services.Configure<IISServerOptions>(options =>
			{
				options.AllowSynchronousIO = true;
			});

			services.AddSingleton(Configuration);
			services.AddMemoryCache();
			services.AddDistributedMemoryCache();
			services.AddSession();

			services.AddControllers().AddNewtonsoftJson(options =>
			{
				options.UseMemberCasing();
			});

			services.AddControllersWithViews();

			services.AddRazorPages(options =>
			{
				options.Conventions.AddPageRoute("/Editor", "/email/edit");
				options.Conventions.AddPageRoute("/Viewer", "/email/view");
			});

			services.AddMvc().AddJsonOptions(options =>
			{
				// don't serialize with CamelCase (see https://github.com/aspnet/Announcements/issues/194)
				options.JsonSerializerOptions.PropertyNamingPolicy = null;
			});
		}

		// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
		public void Configure(
			IApplicationBuilder app,
			IWebHostEnvironment env, 
			ILogger<DefaultExceptionFilterAttribute> exceptionFilterLogger)
		{
            WebHostEnvironment = env;
           
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
            }

			ExceptionFilterLogger = exceptionFilterLogger;

			app.UseCors(
				options => options.WithOrigins("*")
				.AllowAnyMethod()
				.AllowAnyOrigin()
				.AllowAnyHeader()
			);

			app.UseHttpsRedirection();
			app.UseStaticFiles();
			app.UseSession();

			app.UseRouting();

			app.UseEndpoints(endpoints =>
			{
				endpoints.MapControllerRoute(
					name: "default",
					pattern: "{controller=Home}/{action=Index}/{id?}");

				endpoints.MapRazorPages();
			});
		}		
	}
}
