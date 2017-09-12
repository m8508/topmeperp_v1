using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using log4net;
using topmeperp.Models;
using topmeperp.Service;


namespace topmeperp.Controllers
{
    public class CashFlowController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        PurchaseFormService service = new PurchaseFormService();

        // GET: CashFlow 
        [topmeperp.Filter.AuthFilter]
        public ActionResult Index()
        {
            List<CashFlowFunction> lstCashFlow = null;
            lstCashFlow = service.getCashFlow();
            return View(lstCashFlow);
        }

        //取得特定日期收入明細
        public ActionResult CashInFlowItem(string id)
        {
            List<PLAN_ACCOUNT> CashInFlow = null;
            CashInFlow = service.getDebitByDate(id);
            return View(CashInFlow);
        }

        //取得特定日期支出明細
        public ActionResult CashOutFlowItem(string id)
        {
            List<PLAN_ACCOUNT> CashOutFlow = null;
            CashOutFlow = service.getCreditByDate(id);
            return View(CashOutFlow);
        }
    }
}
