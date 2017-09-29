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
            List<ExpenseBudgetSummary> ExpBudget = null;
            if (null != Request["budgetyear"])
            {
                ExpBudget = service.getExpBudgetByYear(int.Parse(Request["budgetyear"]));
            }
            TempData["budgetYear"] = Request["budgetyear"];
            return View(ExpBudget);
        }

        public ActionResult Search()
        {
            List<ExpenseBudgetSummary> ExpBudget = null;
            if (null != Request["budgetyear"])
            {
                ExpBudget = service.getExpBudgetByYear(int.Parse(Request["budgetyear"]));
            }
            TempData["budgetYear"] = Request["budgetyear"];
            return View("ExpenseBudget", ExpBudget);
        }
        /// <summary>
        /// 下載公司費用預算填寫表
        /// </summary>
        public void downLoadExpBudgetForm()
        {
            ExpBudgetFormToExcel poi = new ExpBudgetFormToExcel();
            //檔案位置
            string fileLocation = poi.exportExcel();
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
        //上傳公司費用預算
        [HttpPost]
        public ActionResult uploadExpBudgetTable(HttpPostedFileBase fileBudget)
        {
            int budgetYear = int.Parse(Request["year"]);
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

        //更新公司費用預算
        public String UpdateExpBudget(FormCollection form)
        {
            logger.Info("form:" + form.Count);
            // 先刪除原先資料
            logger.Info("Expense Budget Year =" + form["year"]);
            logger.Info("Delete FIN_EXPENSE_BUDGET By BUDGET_YEAR");
            service.delExpBudgetByYear(int.Parse(form["year"]));
            string msg = "";
            string[] lstsubjctid = form.Get("subjctid").Split(',');
            string[] lst1 = form.Get("janAtm").Split(',');
            string[] lst2 = form.Get("febAtm").Split(',');
            string[] lst3 = form.Get("marAtm").Split(',');
            string[] lst4 = form.Get("aprAtm").Split(',');
            string[] lst5 = form.Get("mayAtm").Split(',');
            string[] lst6 = form.Get("junAtm").Split(',');
            string[] lst7 = form.Get("julAtm").Split(',');
            string[] lst8 = form.Get("augAtm").Split(',');
            string[] lst9 = form.Get("sepAtm").Split(',');
            string[] lst10 = form.Get("octAtm").Split(',');
            string[] lst11 = form.Get("novAtm").Split(',');
            string[] lst12 = form.Get("decAtm").Split(',');
            List<string[]> Atm = new List<string[]>();
            Atm.Add(lst1);
            Atm.Add(lst2);
            Atm.Add(lst3);
            Atm.Add(lst4);
            Atm.Add(lst5);
            Atm.Add(lst6);
            Atm.Add(lst7);
            Atm.Add(lst8);
            Atm.Add(lst9);
            Atm.Add(lst10);
            Atm.Add(lst11);
            Atm.Add(lst12);
            UserService us = new UserService();
            SYS_USER u = (SYS_USER)Session["user"];
            SYS_USER uInfo = us.getUserInfo(u.USER_ID);
            List<FIN_EXPENSE_BUDGET> lst = new List<FIN_EXPENSE_BUDGET>();
            for (int j = 0; j < lstsubjctid.Count(); j++)
            {
                List<FIN_EXPENSE_BUDGET> lstItem = new List<FIN_EXPENSE_BUDGET>();
                for (int i = 0; i < 12; i++)
                {
                    FIN_EXPENSE_BUDGET item = new FIN_EXPENSE_BUDGET();
                    item.BUDGET_YEAR = int.Parse(form["year"]);
                    item.SUBJECT_ID = lstsubjctid[j];
                    item.MODIFY_ID = u.USER_ID;
                    if (Atm[i][j].ToString() == "" && null != Atm[i][j].ToString())
                    {
                        item.AMOUNT = null;
                        item.BUDGET_MONTH = i + 1;
                    }
                    else
                    {
                        item.AMOUNT = decimal.Parse(Atm[i][j]);
                        item.BUDGET_MONTH = i + 1;
                    }
                    item.MODIFY_DATE = DateTime.Now;
                    logger.Info("費用項目代碼 =" + item.SUBJECT_ID + "，and Budget Month = " + item.BUDGET_MONTH + "，and Atm = " + Atm[i][j]);
                    lstItem.Add(item);
                }
                lst.AddRange(lstItem);
            }
            int k = service.refreshExpBudget(lst);
            if (k == 0)
            {
                msg = service.message;
            }
            else
            {
                msg = "更新公司預算費用成功，預算年度為 " + form["year"];
            }

            logger.Info("Request:BUDGET_YEAR =" + form["year"]);
            return msg;
        }

    }
}
