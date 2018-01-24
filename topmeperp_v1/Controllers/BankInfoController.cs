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
                curAmt.CUR_AMOUNT= decimal.Parse(amt[i]);
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
    }
}