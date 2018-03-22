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
        public string statusChange = null;//More 、 Done、Fail
        public string Message = null;
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
        //public void Approve()
        //{

        //}
        //退件
        public void Reject(SYS_USER u, string reason)
        {
            user = u;
            RejectTask(reason);
        }
        //中止
        public void CancelRequest(SYS_USER u)
        {
            user = u;
            logger.Debug("USER :" + user.USER_ID + " Cancel :" + task.task.EXP_FORM_ID);
            CancelRequest();
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
                            if (statusChange == null)
                            {
                                statusChange = "M";//More
                            }
                            return;
                        }
                        else
                        {
                            //No More Step
                            UpdateTask(task.ProcessTask[idx], null);
                            if (statusChange == null)
                            {
                                statusChange = "D";//Done
                            }
                        }
                        break;
                    case "D":
                        //skip task
                        logger.Debug("task id=" + task.ProcessTask[idx].ID);
                        break;
                }
            }
        }
        /// <summary>
        /// 更新現有任務狀態
        /// </summary>
        /// <param name="curTask"></param>
        /// <param name="nextTask"></param>
        protected void UpdateTask(WF_PORCESS_TASK curTask, WF_PORCESS_TASK nextTask)
        {
            string sql4Task = "UPDATE WF_PORCESS_TASK SET STATUS=@status,REMARK=@remark,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE ID=@id";
            string sql4Request = "UPDATE WF_PROCESS_REQUEST SET CURENT_STATE=@state,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                try
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
                    Message = "處理成功";
                }
                catch (Exception ex)
                {
                    statusChange = "F";
                    Message = "處理失敗(" + ex.Message + ")";
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
        }
        /// <summary>
        /// 退件相關狀態處理
        /// </summary>
        protected void RejectTask(string reason)
        {
            string sql4Task = "UPDATE WF_PORCESS_TASK SET STATUS=@status,REMARK=@remark,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            string sql4Request = "UPDATE WF_PROCESS_REQUEST SET CURENT_STATE=@state,MODIFY_USER_ID=@modifyUser,MODIFY_DATE=@modifyDate WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                try
                {
                    //Update Task State roll back to "O"
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("status", "O"));
                    parameters.Add(new SqlParameter("remark", reason));

                    parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("RID", task.task.RID));
                    int i = context.Database.ExecuteSqlCommand(sql4Task, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.RID);

                    //Change Request State
                    parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("state", 1));
                    parameters.Add(new SqlParameter("modifyUser", user.USER_ID));
                    parameters.Add(new SqlParameter("modifyDate", DateTime.Now));
                    parameters.Add(new SqlParameter("RID", task.task.RID));
                    i = context.Database.ExecuteSqlCommand(sql4Request, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.ID);
                    statusChange = "D";
                    Message = "退件作業完成";
                }
                catch (Exception ex)
                {
                    statusChange = "F";
                    Message = "退件作業處理失敗(" + ex.Message + ")";
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
            }
        }
        /// <summary>
        /// 取消表單作業流程
        /// </summary>
        protected void CancelRequest()
        {
            string sql4Task = "DELETE WF_PORCESS_TASK WHERE RID=@RID";
            string sql4Request = "DELETE WF_PROCESS_REQUEST WHERE RID=@RID";
            using (var context = new topmepEntities())
            {
                try
                {
                    //Update Task State roll back to "O"
                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("RID", task.task.RID));
                    int i = context.Database.ExecuteSqlCommand(sql4Task, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.RID);
                    i = context.Database.ExecuteSqlCommand(sql4Request, parameters.ToArray());
                    logger.Debug("i=" + i + "sql" + sql4Task + ",Id" + task.task.ID);
                    statusChange = "D";
                    Message = "表單已取消";
                }
                catch (Exception ex)
                {
                    statusChange = "F";
                    Message = "取消作業處理失敗(" + ex.Message + ")";
                    logger.Error(ex.Message + ":" + ex.StackTrace);
                }
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
    /// <summary>
    /// 公司費用申請控制流程
    /// </summary>
    public class Flow4CompanyExpense : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public string FLOW_KEY = "EXP01";
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"SELECT F.EXP_FORM_ID,F.PROJECT_ID,F.OCCURRED_YEAR,F.OCCURRED_MONTH,F.PAYEE,F.PAYMENT_DATE,F.REMARK REQ_DESC,F.REJECT_DESC REJECT_DESC,
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
        /// <summary>
        /// 取得費用申請清單
        /// </summary>
        /// <param name="occurreddate"></param>
        /// <param name="subjectname"></param>
        /// <param name="expid"></param>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public List<ExpenseFlowTask> getCompanyExpenseRequest(string occurreddate, string subjectname, string expid, string projectid,string status)
        {
            logger.Info("search expense form by " + occurreddate + ", 費用單編號 =" + expid + ", 項目名稱 =" + subjectname + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //申請日期
                if (null!= occurreddate &&  occurreddate != "")
                {
                    sql = sql + " AND F.PAYMENT_DATE=@occurreddate ";
                    parameters.Add(new SqlParameter("occurreddate", occurreddate));
                }
                //申請主旨內容
                if (null != subjectname  && subjectname != "")
                {
                    sql = sql + " AND F.REMARK Like @subjectname ";
                    parameters.Add(new SqlParameter("subjectname", "'%" + subjectname + "%'"));
                }
                //申請單號
                if (null!=expid && expid != "")
                {
                    sql = sql + " AND  F.EXP_FORM_ID Like @expid ";
                    parameters.Add(new SqlParameter("expid", "'%"+ expid + "%'"));
                }
                //工地費用
                if (null != projectid && projectid != "")
                {
                    sql = sql + " AND  F.PROJECT_ID = @projectid ";
                    parameters.Add(new SqlParameter("projectid", projectid));
                }
                //表單狀態
                if (null != status && status != "")
                {
                    sql = sql + " AND  F.STATUS = @status ";
                    parameters.Add(new SqlParameter("status", status));
                }

                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get expense form count=" + lstForm.Count);

            return lstForm;
        }
        /// <summary>
        /// 取的申請程序表單與相關資料
        /// </summary>
        /// <param name="dataKey"></param>
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
            task = new ExpenseTask();
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

        //送審
        public void Send(SYS_USER u, DateTime? paymentdate, string reason)
        {
            //STATUS  0   退件
            //STATUS  10  草稿
            //STATUS  20  審核中
            //STATUS  30  通過
            //STATUS  40  中止
            //
            logger.Debug("CompanyExpenseRequest Send" + task.task.ID);
            base.Send(u);
            if (statusChange != "F")
            {
                int staus = 10;
                if (statusChange == "M")
                {
                    staus = 20;
                }
                else if (statusChange == "D")
                {
                    staus = 30;
                }
                staus = updateForm(paymentdate, reason, staus);
            }
        }
        //更新資料庫資料
        protected int updateForm(DateTime? paymentdate, string reason, int staus)
        {
            string sql = "UPDATE FIN_EXPENSE_FORM SET STATUS=@status,PAYMENT_DATE=@paymentDate,REJECT_DESC=@rejectDesc WHERE EXP_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", staus));
            if (null != paymentdate)
            {
                parameters.Add(new SqlParameter("paymentDate", paymentdate));
            }
            else
            {
                parameters.Add(new SqlParameter("paymentDate", DBNull.Value));
            }

            if (null != reason)
            {
                parameters.Add(new SqlParameter("rejectDesc", reason));
            }
            else
            {
                parameters.Add(new SqlParameter("rejectDesc", DBNull.Value));
            }
            parameters.Add(new SqlParameter("formId", task.task.EXP_FORM_ID));
            using (var context = new topmepEntities())
            {
                logger.Debug("Change CompanyExpenseRequest Status=" + task.task.EXP_FORM_ID + "," + staus);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }

            return staus;
        }

        //退件
        public void Reject(SYS_USER u, DateTime? paymentdate, string reason)
        {
            logger.Debug("CompanyExpenseRequest Reject:" + task.task.RID);
            base.Reject(u, reason);
            if (statusChange != "F")
            {
                updateForm(paymentdate, reason, 0);
            }
        }
        //中止
        public void Cancel(SYS_USER u)
        {
            user = u;
            logger.Info("USER :" + user.USER_ID + " Cancel :" + task.task.EXP_FORM_ID);
            base.CancelRequest(u);
            if (statusChange != "F")
            {
                string sql = "DELETE FIN_EXPENSE_ITEM WHERE EXP_FORM_ID=@formId;DELETE FIN_EXPENSE_FORM WHERE EXP_FORM_ID=@formId;";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", task.task.EXP_FORM_ID));
                using (var context = new topmepEntities())
                {
                    logger.Debug("Cancel CompanyExpenseRequest Status=" + task.task.EXP_FORM_ID);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
    }
    /// <summary>
    /// 工地費用申請
    /// </summary>
    public class Flow4SiteExpense : Flow4CompanyExpense
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //public new string FLOW_KEY = "EXP02";
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"SELECT F.EXP_FORM_ID,F.PROJECT_ID,F.OCCURRED_YEAR,F.OCCURRED_MONTH,F.PAYEE,F.PAYMENT_DATE,F.REMARK REQ_DESC,F.REJECT_DESC REJECT_DESC,
                        R.REQ_USER_ID,R.CURENT_STATE,R.PID,
                        CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
						FROM FIN_EXPENSE_FORM F,WF_PROCESS_REQUEST R,
                        WF_PORCESS_TASK CT ,
		(SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
                        WHERE F.EXP_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
						AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE";
        public Flow4SiteExpense()
        {
            base.FLOW_KEY = "EXP02";
        }
    }

    /// <summary>
    /// 廠商計價請款
    /// </summary>
    public class Flow4Estimation : WorkFlowService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public new string FLOW_KEY = "EST01";
        //處理SQL 預先填入專案代號,設定集合處理參數
        string sql = @"SELECT F.EST_FORM_ID,F.PROJECT_ID,F.CONTRACT_ID,P.SUPPLIER_ID PAYEE,P.FORM_NAME,F.CREATE_DATE PAYMENT_DATE,F.REMARK REQ_DESC,F.REJECT_DESC REJECT_DESC,
                        R.REQ_USER_ID,R.CURENT_STATE,R.PID,
                        CT.* ,M.FORM_URL + METHOD_URL as FORM_URL
						FROM PLAN_ESTIMATION_FORM F,WF_PROCESS_REQUEST R,
                        WF_PORCESS_TASK CT , PLAN_SUP_INQUIRY P,
		(SELECT P.PID,A.SEQ_ID,FORM_URL,METHOD_URL  FROM WF_PROCESS P,WF_PROCESS_ACTIVITY A WHERE P.PID=A.PID ) M
                        WHERE F.EST_FORM_ID= R.DATA_KEY AND R.RID=CT.RID AND R.CURENT_STATE=CT.SEQ_ID
						AND M.PID=R.PID AND M.SEQ_ID=R.CURENT_STATE AND F.CONTRACT_ID = P.INQUIRY_FORM_ID";
        public Flow4Estimation()
        {
            //base.FLOW_KEY = "EST01";
        }
        /// <summary>
        /// 取得廠商計價請款清單
        /// </summary>
        /// <param name="contractid"></param>
        /// <param name="payee"></param>
        /// <param name="estid"></param>
        /// <param name="projectid"></param>
        /// <returns></returns>
        public List<ExpenseFlowTask> getEstimationFormRequest(string contractid, string payee, string estid, string projectid, string status)
        {
            logger.Info("search est form by 合約編號 " + contractid + ", 計價單編號 =" + estid + ", 受款人 =" + payee + ", 專案編號 =" + projectid);
            List<ExpenseFlowTask> lstForm = new List<ExpenseFlowTask>();
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                //合約編號
                if (null != contractid && contractid != "")
                {
                    sql = sql + " AND F.CONTRACT_ID Like @contractid ";
                    parameters.Add(new SqlParameter("contractid", "'%" + contractid + "%'"));
                }
                //廠商名稱(受款人)
                if (null != payee && payee != "")
                {
                    sql = sql + " AND PAYEE Like @payee ";
                    parameters.Add(new SqlParameter("payee", "'%" + payee + "%'"));
                }
                //申請單號
                if (null != estid && estid != "")
                {
                    sql = sql + " AND  F.EST_FORM_ID Like @expid ";
                    parameters.Add(new SqlParameter("estid", "'%" + estid + "%'"));
                }
                //專案名稱
                if (null != projectid && projectid != "")
                {
                    sql = sql + " AND  F.PROJECT_ID = @projectid ";
                    parameters.Add(new SqlParameter("projectid", projectid));
                }
                //表單狀態
                if (null != status && status != "")
                {
                    sql = sql + " AND  F.STATUS = @status ";
                    parameters.Add(new SqlParameter("status", status));
                }

                lstForm = context.Database.SqlQuery<ExpenseFlowTask>(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get est form count=" + lstForm.Count);

            return lstForm;
        }
        /// <summary>
        /// 取的申請程序表單與相關資料
        /// </summary>
        /// <param name="dataKey"></param>
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
            task = new ExpenseTask();
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
        //送審
        public void Send(SYS_USER u, DateTime? paymentdate, string reason)
        {
            //STATUS  0   退件
            //STATUS  10  草稿
            //STATUS  20  審核中
            //STATUS  30  通過
            //STATUS  40  中止
            //
            logger.Debug("EstimationFormRequest Send" + task.task.ID);
            base.Send(u);
            if (statusChange != "F")
            {
                int staus = 10;
                if (statusChange == "M")
                {
                    staus = 20;
                }
                else if (statusChange == "D")
                {
                    staus = 30;
                }
                staus = updateForm(reason, staus);
            }
        }
        //更新資料庫資料
        protected int updateForm(string reason, int staus)
        {
            string sql = "UPDATE PLAN_ESTIMATION_FORM SET STATUS=@status, REJECT_DESC=@rejectDesc WHERE EST_FORM_ID=@formId";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("status", staus));
            
            if (null != reason)
            {
                parameters.Add(new SqlParameter("rejectDesc", reason));
            }
            else
            {
                parameters.Add(new SqlParameter("rejectDesc", DBNull.Value));
            }
            parameters.Add(new SqlParameter("formId", task.task.EST_FORM_ID));
            using (var context = new topmepEntities())
            {
                logger.Debug("Change EstimationFormRequest Status=" + task.task.EST_FORM_ID + "," + staus);
                context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
            }

            return staus;
        }

        //退件
        public void Reject(SYS_USER u, string reason)
        {
            logger.Debug("EstimationFormRequest Reject:" + task.task.RID);
            base.Reject(u, reason);
            if (statusChange != "F")
            {
                updateForm(reason, 0);
            }
        }
        //中止
        public void Cancel(SYS_USER u)
        {
            user = u;
            logger.Info("USER :" + user.USER_ID + " Cancel :" + task.task.EST_FORM_ID);
            base.CancelRequest(u);
            if (statusChange != "F")
            {
                string sql = "DELETE PLAN_ESTIMATION_FORM WHERE EST_FORM_ID=@formId;DELETE PLAN_ESTIMATION_FORM WHERE EST_FORM_ID=@formId;";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("formId", task.task.EST_FORM_ID));
                using (var context = new topmepEntities())
                {
                    logger.Debug("Cancel EstimationFormRequest Status=" + task.task.EST_FORM_ID);
                    context.Database.ExecuteSqlCommand(sql, parameters.ToArray());
                }
            }
        }
    }

}