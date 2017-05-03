using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    public class TypeManageController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        TypeManageService service = new TypeManageService();
        // GET: TypeManage
        public ActionResult Index()
        {
            List<REF_TYPE_MAIN> lst = service.getMainType();
            return View(lst);
        }
        public ActionResult getSubType()
        {
            string typecod1 = Request["typecode1"];
            string typecod2 = Request["typecode2"];
            List<REF_TYPE_SUB> lst = service.getSubType(typecod1.Trim()+ typecod2.Trim());
            logger.Debug("get main type code1=" + typecod1 + ",typecod2=" + typecod2);

            @ViewBag.SearchResult = "取得"+ lst.Count+"筆資料";
            @ViewBag.Typecode = "<a href =\"EditMainType?typecode1=" + typecod1 + "&typecod2=" + typecod2 + "\">編輯</a>";
            return PartialView("_getSubType",lst) ;
        }
        public ActionResult EditMainType()
        {
            string typecod1 = Request["typecode1"];
            string typecod2 = Request["typecode2"];
            logger.Debug("get main type code1=" + typecod1 + ",typecod2=" + typecod2);
            return View();
        }
    }
}