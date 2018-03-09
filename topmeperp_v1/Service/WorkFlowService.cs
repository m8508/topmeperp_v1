using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    //public enum ExecuteState

    //{

    //    Idle,  //停滯

    //    Running, //執行中

    //    Complete, //執行完成

    //    Fail, //錯誤

    //    Cancel, //取消

    //    Jump //跳過此Step

    //}
    //public interface IFlowContext

    //{

    //    System.Collections.Generic.Dictionary Parameters { get; }

    //    T GetParameter(stringkey);

    //    voidSetParameter(stringkey, T value);

    //}
    //public interface IFlowStep
    //{
    //    //當Step執行完畢後觸發
    //    event EventHandler ExecuteComplete;
    //    //當Step執行期間產生未補捉例外時,放置於此處
    //    Exception FailException { get; }
    //    //所有Flow Step共用的Context
    //    IFlowContext Context { get; set; }
    //    //Step Name, 必要欄位
    //    string StepName { get; }
    //    //Step執行狀態
    //    ExecuteState State { get; }
    //    //下一個要執行的Step Name,如未指定則依序執行
    //    string NextFlow { get; set; }
    //    //執行
    //    void Execute();
    //    //確認完成
    //    void Complete();
    //    //要求取消
    //    void Cancel();
    //    //跳過此Step
    //    void Jump();
    //    //Step執行失敗
    //    void Fail(Exception ex);
    //}
    public class WorkFlowService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public WF_PROCESS process { get; set; }
        public List<WF_PROCESS_ACTIVITY> activitys { get; set; }

        public void getFlowAcivities(string flowkey)
        {
            using (var context = new topmepEntities())
            {
                string sql = "SELECT * FROM WF_PROCESS WHERE PROCESS_CODE=@processCode";
                logger.Debug("get process PROCESS_CODE=" + flowkey + ",sql=" + sql);
                process = context.WF_PROCESS.SqlQuery(sql, new SqlParameter("processCode", flowkey)).First();
                sql = "SELECT * FROM WF_PROCESS_ACTIVITY WHERE PID=@pid";
                logger.Debug("get activities pid=" + process.PID + ",sql=" + sql);
                activitys = context.WF_PROCESS_ACTIVITY.SqlQuery(sql, new SqlParameter("pid", process.PID)).ToList();
            }
        }
    }
    public class Flow4CompanyExpense : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static string FLOW_KEY = "EXP01";
        //對應到表單的請求資料
        public WF_PROCESS_REQUEST ProcessRequest = null;
        //表單簽核所需的步驟
        public List<WF_PORCESS_TASK> ProcessTask = null;

        public Flow4CompanyExpense()
        {

        }
        public List<ExpenseFlowTask> getCompanyExpenseRequest(string occurreddate, string subjectname, string expid, string projectid)
        {
            logger.Info("search expense form by " + occurreddate + ", 費用單編號 =" + expid + ", 項目名稱 =" + subjectname + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = @"SELECT F.EXP_FORM_ID,F.PROJECT_ID,F.OCCURRED_YEAR,F.OCCURRED_MONTH,F.PAYEE,F.PAYMENT_DATE,F.REMARK REQ_DESC,
                        R.REQ_USER_ID,R.CURENT_STATE,R.PID,
                        CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
						FROM FIN_EXPENSE_FORM F,WF_PROCESS_REQUEST R,
                        WF_PORCESS_TASK CT ,
		(SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
                        WHERE F.EXP_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
						AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE";

            using (var context = new topmepEntities())
            {
                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("id", projectid)).ToList();
            }
            logger.Info("get expense form count=" + lstForm.Count);

            return lstForm;
        }
        public void getProcessTask(string dataId)
        {

        }
        //任務狀態變化

        //表單送審後，啟動對應程序
        public void iniRequest(SYS_USER u, string DataKey)
        {
            getFlowAcivities(FLOW_KEY);
            //建立表單追蹤資料Index
            ProcessRequest = new WF_PROCESS_REQUEST();
            ProcessRequest.PID = process.PID;
            ProcessRequest.REQ_USER_ID = u.USER_ID;
            ProcessRequest.CREATE_DATE = DateTime.Now;
            ProcessRequest.DATA_KEY = DataKey;
            ProcessRequest.CURENT_STATE = 1;
            //建立簽核任務追蹤要項
            ProcessTask = new List<WF_PORCESS_TASK>();
            foreach (WF_PROCESS_ACTIVITY activity in activitys)
            {
                WF_PORCESS_TASK t = new WF_PORCESS_TASK();
                t.ACTIVITY_ID = activity.ID;
                t.CREATE_DATE = DateTime.Now;
                t.CREATE_USER_ID = u.USER_ID;
                t.NOTE = activity.ACTIVITY_NAME;
                t.STATUS = "O";//參考TaskStatus
                t.SEQ_ID = activity.SEQ_ID;
                t.DEP_CODE = activity.DEP_CODE;
                ProcessTask.Add(t);
            }
            using (var context = new topmepEntities())
            {
                context.WF_PROCESS_REQUEST.Add(ProcessRequest);
                int i = context.SaveChanges();
                foreach (WF_PORCESS_TASK t in ProcessTask)
                {
                    //t.ID = DBNull.Value;
                    t.RID = ProcessRequest.RID;
                }
                context.WF_PORCESS_TASK.AddRange(ProcessTask);
                i = context.SaveChanges();
                logger.Debug("Create Task Records =" + i);
            }
        }

        public enum TaskStatus
        {
            O, // Open
            R, //Running
            C, //Complete執行完成
            RJ, //Reject 拒絕 
            CAN, //Cancel 取消
            Jump //跳過此Step
        }
    }
}