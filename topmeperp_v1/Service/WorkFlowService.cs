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

        public ExpenseTask task = null;
        protected string sql4Request = "SELECT * FROM WF_PROCESS_REQUEST WHERE DATA_KEY=@dataKey";
        protected string sql4Task = "SELECT * from WF_PORCESS_TASK WHERE RID=@rid ORDER BY SEQ_ID";
        protected SYS_USER user = null;
        public enum TaskStatus
        {
            O, // Open
            R, //Running
            C, //Complete執行完成
            RJ, //Reject 拒絕 
            CAN, //Cancel 取消
            Jump //跳過此Step
        }

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
        //送審
        public void Send(SYS_USER u)
        {
            user = u;
            processTask();

        }
        //審核通過
        public void Approve()
        {

        }
        //退件
        public void Reject()
        {

        }
        //中止
        public void Cancel()
        {

        }
        protected void processTask()
        {
            int idx = 0;
            for (idx = 0; idx < task.ProcessTask.Count; idx++)
            {
                switch (task.ProcessTask[idx].STATUS)
                {
                    case "O":
                        //change request status
                        if (idx + 1 < task.ProcessTask.Count)
                        {
                            //Has Next Step
                            UpdateTask(task.ProcessTask[idx], task.ProcessTask[idx + 1]);
                            return;
                        }
                        else
                        {
                            //No More Step
                            UpdateTask(task.ProcessTask[idx], null);
                        }
                        break;
                    case "D":
                        //skip task
                        logger.Debug("task id=" + task.ProcessTask[idx].ID);
                        break;
                }
            }
        }
        protected void UpdateTask(WF_PORCESS_TASK curTask, WF_PORCESS_TASK nextTask)
        {
            string sql4Task = "UPDATE WF_PORCESS_TASK SET STATUS=@status,REMARK=@remark,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE ID=@id";
            string sql4Request = "UPDATE WF_PROCESS_REQUEST SET CURENT_STATE=@state,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                //Update Task State=Done
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("status", "D"));
                if (null != curTask.REMARK)
                {
                    parameters.Add(new SqlParameter("remark", curTask.REMARK));
                }
                else
                {
                    parameters.Add(new SqlParameter("remark", DBNull.Value));
                }
                parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                parameters.Add(new SqlParameter("id", curTask.ID));
                //Change Request State
                int i = context.Database.ExecuteSqlCommand(sql4Task, parameters.ToArray());
                logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + curTask.ID);
                parameters = new List<SqlParameter>();
                if (null != nextTask)
                {
                    parameters.Add(new SqlParameter("state", nextTask.SEQ_ID));
                }
                else
                {
                    parameters.Add(new SqlParameter("state", -1));
                }
                parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                parameters.Add(new SqlParameter("RID", curTask.RID));

                i = context.Database.ExecuteSqlCommand(sql4Request, parameters.ToArray());
                logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + curTask.ID);
            }
        }
        //取得表單與對應的流程資料
        public void getRequest(string datakey)
        {
            using (var context = new topmepEntities())
            {
                if (task == null)
                {
                    logger.Warn("task is null");
                    task = new ExpenseTask();
                }
                logger.Debug("get Request =" + datakey + ",sql=" + sql4Request);
                task.ProcessRequest = context.WF_PROCESS_REQUEST.SqlQuery(sql4Request, new SqlParameter("dataKey", datakey)).First();
                logger.Debug("get task rid=" + task.ProcessRequest.RID + ",sql=" + sql4Task);
                task.ProcessTask = context.WF_PORCESS_TASK.SqlQuery(sql4Task, new SqlParameter("rid", task.ProcessRequest.RID)).ToList();
            }
        }
    }
    public class Flow4CompanyExpense : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static string FLOW_KEY = "EXP01";
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"SELECT F.EXP_FORM_ID,F.PROJECT_ID,F.OCCURRED_YEAR,F.OCCURRED_MONTH,F.PAYEE,F.PAYMENT_DATE,F.REMARK REQ_DESC,
                        R.REQ_USER_ID,R.CURENT_STATE,R.PID,
                        CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
						FROM FIN_EXPENSE_FORM F,WF_PROCESS_REQUEST R,
                        WF_PORCESS_TASK CT ,
		(SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
                        WHERE F.EXP_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
						AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE";

        public Flow4CompanyExpense()
        {

        }
        public List<ExpenseFlowTask> getCompanyExpenseRequest(string occurreddate, string subjectname, string expid, string projectid)
        {
            logger.Info("search expense form by " + occurreddate + ", 費用單編號 =" + expid + ", 項目名稱 =" + subjectname + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            using (var context = new topmepEntities())
            {
                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("id", projectid)).ToList();
            }
            logger.Info("get expense form count=" + lstForm.Count);

            return lstForm;
        }
        public void getTask(string dataKey)
        {
            task = new ExpenseTask();
            sql = sql + " AND R.DATA_KEY=@datakey";
            using (var context = new topmepEntities())
            {
                try
                {
                    task.task = context.Database.SqlQuery<ExpenseFlowTask>(sql, new SqlParameter("datakey", dataKey)).First();
                }
                catch (Exception ex)
                {
                    logger.Warn("not task!! ex=" + ex.Message + "," + ex.StackTrace);
                }
            }
        }

        //表單送審後，啟動對應程序
        public void iniRequest(SYS_USER u, string DataKey)
        {
            getFlowAcivities(FLOW_KEY);
            //建立表單追蹤資料Index
            task.ProcessRequest = new WF_PROCESS_REQUEST();
            task.ProcessRequest.PID = process.PID;
            task.ProcessRequest.REQ_USER_ID = u.USER_ID;
            task.ProcessRequest.CREATE_DATE = DateTime.Now;
            task.ProcessRequest.DATA_KEY = DataKey;
            task.ProcessRequest.CURENT_STATE = 1;
            //建立簽核任務追蹤要項
            task.ProcessTask = new List<WF_PORCESS_TASK>();
            foreach (WF_PROCESS_ACTIVITY activity in activitys)
            {
                WF_PORCESS_TASK t = new WF_PORCESS_TASK();
                t.ACTIVITY_TYPE = activity.ACTIVITY_TYPE;
                t.CREATE_DATE = DateTime.Now;
                t.CREATE_USER_ID = u.USER_ID;
                t.NOTE = activity.ACTIVITY_NAME;
                t.STATUS = "O";//參考TaskStatus
                t.SEQ_ID = activity.SEQ_ID;
                t.DEP_CODE = activity.DEP_CODE;
                task.ProcessTask.Add(t);
            }
            using (var context = new topmepEntities())
            {
                context.WF_PROCESS_REQUEST.Add(task.ProcessRequest);
                int i = context.SaveChanges();
                foreach (WF_PORCESS_TASK t in task.ProcessTask)
                {
                    //t.ID = DBNull.Value;
                    t.RID = task.ProcessRequest.RID;
                }
                context.WF_PORCESS_TASK.AddRange(task.ProcessTask);
                i = context.SaveChanges();
                logger.Debug("Create Task Records =" + i);
            }
        }

    }
}