using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(topmeperp_v1.Startup))]
namespace topmeperp_v1
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
