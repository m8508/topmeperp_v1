using log4net;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
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
        public List<BankLoanInfoExt> getAllBankLoan()
        {
            List<BankLoanInfoExt> lstBankLoan = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    string sql = "SELECT * , (SELECT SUM(TRANSACTION_TYPE * AMOUNT) FROM FIN_LOAN_TRANACTION T WHERE T.BL_ID = B.BL_ID) SumTransactionAmount  FROM FIN_BANK_LOAN B";

                    lstBankLoan = context.Database.SqlQuery<BankLoanInfoExt>(sql).ToList();
                    logger.Info("new bank loan records=" + lstBankLoan.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return lstBankLoan;
        }
        //取得貸款帳戶交易資料
        public BankLoanInfo getBankLoan(string bl_id)
        {
            BankLoanInfo item = new BankLoanInfo();
            using (var context = new topmepEntities())
            {
                try
                {
                    logger.Info("get bank transaction BL_ID=" + bl_id);
                    item.LoanInfo = context.FIN_BANK_LOAN.Find(long.Parse(bl_id));
                    long blid = long.Parse(bl_id);
                    item.LoanTransaction = context.FIN_LOAN_TRANACTION.Where(b => b.BL_ID == blid).ToList();
                    //取得期數與匯總金額
                    string sql = "SELECT MAX(ISNULL(PERIOD,0)) CUR_PERIOD,SUM(TRANSACTION_TYPE*AMOUNT) AMOUNT  from FIN_LOAN_TRANACTION WHERE BL_ID=@BL_ID";
                    Dictionary<string, object> para = new Dictionary<string, object>();
                    para.Add("BL_ID", blid);
                    DataSet ds =ExecuteStoreQuery(sql, System.Data.CommandType.Text, para);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        item.CurPeriod = long.Parse(ds.Tables[0].Rows[0]["CUR_PERIOD"].ToString());
                        item.SumTransactionAmount = (decimal)ds.Tables[0].Rows[0]["AMOUNT"];
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return item;
        }
        //新增貸款帳戶資料
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
        /// <summary>
        /// 增加借款還款紀錄
        /// </summary>
        /// <param name="loanTransaction"></param>
        /// <returns></returns>
        public int addBankLoanTransaction(List<FIN_LOAN_TRANACTION> loanTransaction)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.FIN_LOAN_TRANACTION.AddRange(loanTransaction);
                    i = context.SaveChanges();
                    logger.Info("new bank loan transaction record=" + loanTransaction.Count);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message + ":StackTrace=" + ex.StackTrace);
                }
            }
            return i;
        }
    }
}
