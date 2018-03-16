using Microsoft.AspNet.Identity;
using Microsoft.Owin;
using Microsoft.Owin.Security.Cookies;
using Owin;
using log4net;

namespace topmeperp
{
    public partial class Startup
    {
        ILog log = log4net.LogManager.GetLogger(typeof(Startup));
        // 如需設定驗證的詳細資訊，請瀏覽 http://go.microsoft.com/fwlink/?LinkId=301864
        public void ConfigureAuth(IAppBuilder app)
        {
            // 配置Middleware 組件
            log.Info("UseCookieAuthentication");
            app.UseCookieAuthentication(new CookieAuthenticationOptions
            {
                AuthenticationType = DefaultAuthenticationTypes.ApplicationCookie,
                LoginPath = new PathString("/Index/Login"),
                CookieSecure = CookieSecureOption.Never,
            });
        }
    }
    public class AppInfo
    {
        public static string Version="1.0.27";
    }
}