using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(WebReport.Startup))]
namespace WebReport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
