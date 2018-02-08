using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    //銀行帳戶服務層
    public class ContextService4BankInfo : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public ContextService4BankInfo()
        {

        }
        //增加銀行帳戶
        public int addBankInfo(FIN_BANK_ACCOUNT account)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_BANK_ACCOUNT.Add(account);
                    i = context.SaveChanges();
                    logger.Info("new bank account record=" + account.BANK_ACCOUNT_ID + ",initial amount=" + account.CUR_AMOUNT + ",curDate=" + account.CUR_DATE);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            logger.Info("add bankAccount count =" + i);
            return i;
        }
        //取得銀行帳戶
        public List<FIN_BANK_ACCOUNT> getAllBankAccount()
        {
            List<FIN_BANK_ACCOUNT> lstBankAccount = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    lstBankAccount = context.FIN_BANK_ACCOUNT.ToList();
                    logger.Info("new bank account records=" + lstBankAccount.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return lstBankAccount;
        }
        //更新帳戶現額資料
        public void updateBankAccount(List<FIN_BANK_ACCOUNT> lstBankAccount)
        {
            using (var context = new topmepEntities())
            {
                string sql = "UPDATE FIN_BANK_ACCOUNT SET CUR_AMOUNT=@CurAmt,CUR_DATE=@CurDate,MODIFY_ID=@ModifyId,MODIFY_DATE=@ModifyDate WHERE BANK_ACCOUNT_ID=@BankAccountId";
                foreach (FIN_BANK_ACCOUNT account in lstBankAccount)
                {
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("BankAccountId", account.BANK_ACCOUNT_ID));
                    parameters.Add(new SqlParameter("CurAmt", account.CUR_AMOUNT));
                    parameters.Add(new SqlParameter("CurDate", account.CUR_DATE));
                    parameters.Add(new SqlParameter("ModifyId", account.MODIFY_ID));
                    parameters.Add(new SqlParameter("ModifyDate", account.MODIFY_DATE));
                    logger.Info("update bank account record=" + account.BANK_ACCOUNT_ID + ",initial amount=" + account.CUR_AMOUNT + ",curDate=" + account.CUR_DATE);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
        //取得貸款帳戶資料
        public List<FIN_BANK_LOAN> getAllBankLoan()
        {
            List<FIN_BANK_LOAN> lstBankLoan = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    lstBankLoan = context.FIN_BANK_LOAN.ToList();
                    logger.Info("new bank loan records=" + lstBankLoan.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return lstBankLoan;
        }
        //取得貸款帳戶資料
        public int addBankLoan(FIN_BANK_LOAN bankloan)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_BANK_LOAN.Add(bankloan);
                    i = context.SaveChanges();
                    logger.Info("new bank loan record=" + bankloan.ACCOUNT_NAME);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            logger.Info("add bankloan count =" + i);
            return i;
        }
    }
}
