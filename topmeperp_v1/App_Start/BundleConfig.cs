using System.Web;
using System.Web.Optimization;

namespace topmeperp
{
    public class BundleConfig
    {
        // 如需「搭配」的詳細資訊，請瀏覽 http://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Scripts/jquery.validate*"));

            //jquery ui by ph
            bundles.Add(new ScriptBundle("~/bundles/jqueryui").Include(
                        "~/Scripts/jquery-ui-{version}.js"));

            // 使用開發版本的 Modernizr 進行開發並學習。然後，當您
            // 準備好實際執行時，請使用 http://modernizr.com 上的建置工具，只選擇您需要的測試。
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            //modify datatime picker!!
            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                      "~/Scripts/bootstrap.js",
                      "~/Scripts/moment-with-locales.js", //new
                      "~/Scripts/bootstrap-datetimepicker.js", //new
                      "~/Scripts/respond.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/bootstrap.css",
                      "~/Content/bootstrap-datetimepicker.css", //new
                      "~/Content/site.css"));

            //jquery ui css by ph
            bundles.Add(new StyleBundle("~/Content/themes/base/css").Include(
              "~/Content/themes/base/jquery.ui.core.css",
              "~/Content/themes/base/jquery.ui.resizable.css",
              "~/Content/themes/base/jquery.ui.selectable.css",
              "~/Content/themes/base/jquery.ui.accordion.css",
              "~/Content/themes/base/jquery.ui.autocomplete.css",
              "~/Content/themes/base/jquery.ui.button.css",
              "~/Content/themes/base/jquery.ui.dialog.css",
              "~/Content/themes/base/jquery.ui.slider.css",
              "~/Content/themes/base/jquery.ui.tabs.css",
              "~/Content/themes/base/jquery.ui.datepicker.css",
              "~/Content/themes/base/jquery.ui.progressbar.css",
              "~/Content/themes/base/jquery.ui.theme.css"));

            //jquery ui tutorial
            //bundles.Add(new StyleBundle("~/Content/themes/base/css").Include(
            //            "~/Content/themes/base/jquery.ui.core.css",
             ///           "~/Content/themes/base/jquery.ui.autocomplete.css",
            //            "~/Content/themes/base/jquery.ui.theme.css"));
        }
    }
}
