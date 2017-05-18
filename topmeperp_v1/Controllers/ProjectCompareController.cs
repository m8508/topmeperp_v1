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
    public class ProjectCompareController : Controller
    {
        RptCompareProjectPrice service = new RptCompareProjectPrice();
        static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: ProjectCompare
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult getCompareData(FormCollection f)
        {
            log.Info("Source ProijectID=" + f["srcprojectid"] + ",Target ProjectId=" + f["tarprojectid"] + "," + f["hasPrice"]);
            bool hasPriec = false;
            bool hasProject = false;
            if (null != f["hasPrice"])
            {
                hasPriec = true;
            }
            if (null != f["hasProject"])
            {
                hasProject = true;
            }
             
            List<ProjectCompareData> lst = service.RtpGetPriceFromExistProject(f["srcprojectid"], f["tarprojectid"], hasProject, hasPriec);
            ViewBag.Result = "共取得" + lst.Count + "筆資料!!";
            return PartialView("_CompareData", lst);
        }
        public void Update(FormCollection f)
        {
            string[] lstItem = f["chkItem"].Split(',');
            for(int i=0; i<lstItem.Count(); i++)
            {
                log.Info("ITEM_INFO=" + lstItem[i]);
            }
        }
    }
}