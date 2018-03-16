﻿using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;

namespace topmeperp.Controllers
{
    /// <summary>
    /// 銀行存款與貸/還款交易記錄
    /// </summary>
    public class BankInfoController : Controller
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // GET: BankInfo
        public ActionResult Index()
        {
            SYS_USER u = (SYS_USER)Session["user"];
            logger.Debug(u.USER_ID + " Query BankInfo!!");
            ContextService4BankInfo service = new ContextService4BankInfo();
            List<FIN_BANK_ACCOUNT> lstBankAccount = service.getAllBankAccount();
            return View(lstBankAccount);
        }
        /// <summary>
        /// 增加銀行帳戶資料
        /// </summary>
        /// <returns></returns>
        public string addBankAccount(FormCollection f)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            FIN_BANK_ACCOUNT bankaccount = new FIN_BANK_ACCOUNT();
            bankaccount.BANK_ID = Request["BANK_ID"];
            bankaccount.BANK_NAME = Request["BANK_NAME"];
            bankaccount.BRANCH_NAME = Request["BRANCH_NAME"];
            bankaccount.ACCOUNT_NAME = Request["ACCOUNT_NAME"];
            bankaccount.ACCOUNT_NO = Request["ACCOUNT_NO"];
            decimal curAmount = decimal.Parse(Request["CUR_AMOUNT"]);
            bankaccount.CUR_AMOUNT = curAmount;
            bankaccount.CUR_DATE = DateTime.Parse(Request["CUR_DATE"]);
            bankaccount.CREATE_ID = u.USER_ID;
            bankaccount.CREATE_DATE = DateTime.Now;
            ContextService4BankInfo service = new ContextService4BankInfo();
            int i = service.addBankInfo(bankaccount);
            if (i > 0)
            {
                return "更新成功(" + bankaccount.BANK_ACCOUNT_ID + ")!!";
            }
            else
            {
                return "更新失敗!!";
            }
        }
        /// <summary>
        /// 修改銀行帳戶現額資料
        /// </summary>
        /// <returns></returns>
        public void updateBankAccount(FormCollection f)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            string[] acctId = f["BankAccountId"].Split(',');
            //可處理千分位符號!!
            string[] amt = (string[])f.GetValue("curAmount").RawValue;
            string[] curDate = f["curDate"].Split(',');
            List<FIN_BANK_ACCOUNT> lstCurAmt = new List<FIN_BANK_ACCOUNT>();
            for (int i = 0; i < amt.Length; i++)
            {
                FIN_BANK_ACCOUNT curAmt = new FIN_BANK_ACCOUNT();
                curAmt.BANK_ACCOUNT_ID = long.Parse(acctId[i]);
                curAmt.CUR_AMOUNT = decimal.Parse(amt[i]);
                curAmt.CUR_DATE = DateTime.Parse(curDate[i]);
                curAmt.MODIFY_ID = u.USER_ID;
                curAmt.MODIFY_DATE = DateTime.Now;
                lstCurAmt.Add(curAmt);
            }
            logger.Info("Modify Records=" + lstCurAmt.Count);
            ContextService4BankInfo service = new ContextService4BankInfo();
            service.updateBankAccount(lstCurAmt);
            Response.Redirect("Index");
        }
        /// <summary>
        /// 取得貸款銀行帳戶資料
        /// </summary>
        /// <returns></returns>
        public ActionResult BankLoanList()
        {
            ContextService4BankInfo service = new ContextService4BankInfo();
            List<BankLoanInfoExt> lstBankLoan = service.getAllBankLoan();
            service.getAllPlan();

            ViewData.Add("plans", service.tndProjectModels.planList);
            if (service.tndProjectModels.planList != null)
            {
                SelectList plans = new SelectList(service.tndProjectModels.planList, "PROJECT_ID", "PROJECT_NAME");

                ViewBag.plans = plans;
                //將資料存入TempData 減少不斷讀取資料庫
                TempData.Remove("plans");
                TempData.Add("plans", service.tndProjectModels.planList);
            }
            return View(lstBankLoan);
        }
        /// <summary>
        /// 增加貸款銀行基本資料
        /// </summary>
        /// <returns></returns>
        public string addBankLoan(FormCollection f)
        {
            SYS_USER u = (SYS_USER)Session["user"];
            FIN_BANK_LOAN bankloanInfo = new FIN_BANK_LOAN();
            bankloanInfo.BANK_ID = Request["BANK_ID"];
            bankloanInfo.BANK_NAME = Request["BANK_NAME"];
            //bankloanInfo.BRANCH_NAME = Request["BRANCH_NAME"];
            bankloanInfo.ACCOUNT_NAME = Request["ACCOUNT_NAME"];
            //bankloanInfo.ACCOUNT_NO = Request["ACCOUNT_NO"];

            bankloanInfo.START_DATE = DateTime.Parse(Request["START_DATE"]);
            bankloanInfo.DUE_DATE = DateTime.Parse(Request["DUE_DATE"]);

            bankloanInfo.PERIOD_COUNT = int.Parse(Request["PERIOD_COUNT"]);
            decimal quota = decimal.Parse(Request["QUOTA"]);
            bankloanInfo.QUOTA = quota;
            if (Request["AR_PAYBACK_RATIO"] != "")
            {
                bankloanInfo.AR_PAYBACK_RATIO = decimal.Parse(Request["AR_PAYBACK_RATIO"]);
            }
            if (Request["CUM_AR_RATIO"] != "")
            {
                bankloanInfo.CUM_AR_RATIO = decimal.Parse(Request["CUM_AR_RATIO"]);
            }
            //if (Request["QUOTA_AVAILABLE_RATIO"] != "")
            //{
                //bankloanInfo.QUOTA_AVAILABLE_RATIO = decimal.Parse(Request["QUOTA_AVAILABLE_RATIO"]);
            //}
            //else
            //{
                //bankloanInfo.QUOTA_AVAILABLE_RATIO = 100;
            //}
            //bankloanInfo.QUOTA_RECYCLABLE = Request["QUOTA_RECYCLABLE"];
            bankloanInfo.REMARK = Request["REMARK"];
            bankloanInfo.ACCOUNT_NAME = Request["ACCOUNT_NAME"];
            if (Request["plans"] != "")
            {
                bankloanInfo.PROJECT_ID = Request["plans"];
            }
            bankloanInfo.CREATE_ID = u.USER_ID;
            bankloanInfo.CREATE_DATE = DateTime.Now;
            ContextService4BankInfo service = new ContextService4BankInfo();
            int i = service.addBankLoan(bankloanInfo);
            if (i > 0)
            {
                return "更新成功(" + bankloanInfo.BL_ID + ")!!";
            }
            else
            {
                return "更新失敗!!";
            }
        }
        /// <summary>
        /// 貸還款交易維護
        /// </summary>
        /// <returns></returns>
        public ActionResult BankLoanTransaction()
        {
            string blid = Request["BL_ID"];
            ContextService4BankInfo service = new ContextService4BankInfo();
            BankLoanInfo loanInfo = service.getBankLoan(blid);
            logger.Debug("Bank Loan Transaction:" + blid);
            return View(loanInfo);
        }
        /// <summary>
        /// 增加借還款紀錄
        /// </summary>
        public void addBankLoanTransaction()
        {
            logger.Info("bl_id=" + Request["bl_id"]);
            long blid = long.Parse(Request["bl_id"]);
            int period = int.Parse(Request["period"]);
            string[] formKeys = Request.Form.AllKeys;
            List<FIN_LOAN_TRANACTION> lstLoanTransaction = new List<FIN_LOAN_TRANACTION>();
            List<string> lstFormKey = new List<string>();
            for (int i = 0; i < formKeys.Length; i++)
            {
                logger.Debug("key=" + formKeys[i]);
                if (formKeys[i].StartsWith("EVENT_DATE"))
                {
                    lstFormKey.Add(formKeys[i]);
                }
            }
            //處理借款紀錄
            SYS_USER u = (SYS_USER)Session["user"];
            foreach (string key in lstFormKey)
            {
                string[] keyAry = key.Split('.');
                //借款記錄
                if (Request["LOAN_AMOUNT." + keyAry[1]].Trim() != "")
                {
                    FIN_LOAN_TRANACTION loanTransaction = new FIN_LOAN_TRANACTION();
                    loanTransaction.BL_ID = blid;
                    loanTransaction.PERIOD = period + 1;
                    loanTransaction.EVENT_DATE = DateTime.Parse(Request["EVENT_DATE." + keyAry[1]]);
                    loanTransaction.TRANSACTION_TYPE = -1;
                    loanTransaction.AMOUNT = decimal.Parse(Request["LOAN_AMOUNT." + keyAry[1]]);
                    loanTransaction.REMARK = Request["REMARK." + keyAry[1]].Trim();
                    logger.Info("Event Date=" + loanTransaction.EVENT_DATE + ",Loan Amount=" + loanTransaction.AMOUNT);

                    loanTransaction.CREATE_ID = u.USER_ID;
                    loanTransaction.CREATE_DATE = DateTime.Now;
                    lstLoanTransaction.Add(loanTransaction);
                    period++;
                }
                //還款記錄
                if (Request["PAYBACK_LOAN_AMOUNT." + keyAry[1]].Trim() != "")
                {
                    FIN_LOAN_TRANACTION loanTransaction = new FIN_LOAN_TRANACTION();
                    loanTransaction.BL_ID = blid;
                    loanTransaction.PERIOD = period + 1;
                    loanTransaction.PAYBACK_DATE = DateTime.Parse(Request["EVENT_DATE." + keyAry[1]]);
                    loanTransaction.TRANSACTION_TYPE = 1;
                    loanTransaction.AMOUNT = decimal.Parse(Request["PAYBACK_LOAN_AMOUNT." + keyAry[1]]);
                    loanTransaction.REMARK = Request["REMARK." + keyAry[1]].Trim();
                    logger.Info("PAYBACK_DATE=" + loanTransaction.PAYBACK_DATE + ",PayBack Amount=" + loanTransaction.AMOUNT);

                    loanTransaction.CREATE_ID = u.USER_ID;
                    loanTransaction.CREATE_DATE = DateTime.Now;
                    lstLoanTransaction.Add(loanTransaction);
                    period++;
                }
                if (lstLoanTransaction.Count > 0)
                {
                    ContextService4BankInfo service = new ContextService4BankInfo();
                    service.addBankLoanTransaction(lstLoanTransaction);
                }
            }

            Response.Redirect("/BankInfo/BankLoanTransaction?BL_ID=" + Request["bl_id"]);
        }
    }
}