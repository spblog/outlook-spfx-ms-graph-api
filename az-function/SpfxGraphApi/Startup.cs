using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SpfxGraphApi;

[assembly: FunctionsStartup(typeof(Startup))]

namespace SpfxGraphApi
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var appInfo = new AppInfo();
            config.Bind(appInfo);

            builder.Services.AddSingleton(appInfo);
        }
    }
}