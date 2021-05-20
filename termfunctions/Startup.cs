using System.Security.Cryptography.X509Certificates;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth;
using PnP.Core.Services.Builder.Configuration;

[assembly: FunctionsStartup(typeof(mloitzl.sharepoint.functions.Startup))]

namespace mloitzl.sharepoint.functions
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var settings = new Settings();
            config.Bind(settings);

            builder.Services.AddPnPCore(options =>
            {
                options.DisableTelemetry = true;

                var authProvider = new X509CertificateAuthenticationProvider(
                    settings.ClientId.ToString(),
                    settings.TenantId.ToString(),
                    StoreName.My,
                    StoreLocation.CurrentUser,
                    settings.CertificateThumbprint);
                options.DefaultAuthenticationProvider = authProvider;

                options.Sites.Add("Default",
                       new PnPCoreSiteOptions
                       {
                           SiteUrl = settings.SiteUrl.ToString(),
                           AuthenticationProvider = authProvider
                       });
            });
        }
    }
}