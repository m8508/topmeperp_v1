using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;


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

        public ActionResult ExpenseBudget()
        {
            logger.Info("Access to Expense Budget Page !!");
            //取得已寫入之公司費用預算資料
            return View();
        }

        [HttpPost]
        public ActionResult ExpenseBudget(int year)
        {
            logger.Info("Get Expense Budget by Budget Year =" + year);
            //取得已寫入之公司費用預算資料
            List<ExpenseBudgetSummary> ExpBudget = service.getExpBudgetByYear(year);
            return View(ExpBudget);
        }
        /// <summary>
        /// 下載公司費用預算填寫表
        /// </summary>
        public void downLoadExpBudgetForm()
        {
            int budgetYear = int.Parse(Request["budgetyear"]);
            ExpBudgetFormToExcel poi = new ExpBudgetFormToExcel();
            //檔案位置
            string fileLocation = poi.exportExcel(budgetYear);
            //檔案名稱 HttpUtility.UrlEncode預設會以UTF8的編碼系統進行QP(Quoted-Printable)編碼，可以直接顯示的7 Bit字元(ASCII)就不用特別轉換。
            string filename = HttpUtility.UrlEncode(Path.GetFileName(fileLocation));
            Response.Clear();
            Response.Charset = "utf-8";
            Response.ContentType = "text/xls";
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
            ///"\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_ID + ".xlsx"
            Response.WriteFile(fileLocation);
            Response.End();
        }
        //上傳預算
        [HttpPost]
        public ActionResult uploadExpBudgetTable(HttpPostedFileBase fileBudget)
        {
            int budgetYear = int.Parse(Request["budgetyear"]);
            logger.Info("Upload Expense Budget Table for budget year =" + budgetYear);
            string message = "";

            if (null != fileBudget && fileBudget.ContentLength != 0)
            {
                //2.解析Excel
                logger.Info("Parser Excel data:" + fileBudget.FileName);
                //2.1 設定Excel 檔案名稱
                var fileName = Path.GetFileName(fileBudget.FileName);
                var path = Path.Combine(ContextService.strUploadPath, fileName);
                logger.Info("save excel file:" + path);
                fileBudget.SaveAs(path);
                //2.2 開啟Excel 檔案
                logger.Info("Parser Excel File Begin:" + fileBudget.FileName);
                ExpBudgetFormToExcel budgetservice = new ExpBudgetFormToExcel();
                budgetservice.InitializeWorkbook(path);
                //解析預算數量
                List<FIN_EXPENSE_BUDGET> lstExpBudget = budgetservice.ConvertDataForExpBudget(budgetYear);
                //2.3 記錄錯誤訊息
                message = budgetservice.errorMessage;
                //2.4
                logger.Info("Delete FIN_EXPENSE_BUDGET By Year");
                service.delExpBudgetByYear(budgetYear);
                message = message + "<br/>舊有資料刪除成功 !!";
                //2.5 
                logger.Info("Add All FIN_EXPENSE_BUDGET to DB");
                service.refreshExpBudget(lstExpBudget);
                message = message + "<br/>資料匯入完成 !!";
            }
            TempData["result"] = message;
            return RedirectToAction("ExpenseBudget");
        }
    }
}
