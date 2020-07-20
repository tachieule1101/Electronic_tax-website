using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(VietHoaLayout.Startup))]
namespace VietHoaLayout
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
