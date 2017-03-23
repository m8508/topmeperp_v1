using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class ContextService
    {
        public topmepEntities db;// = new topmepEntities();
        //定義上傳檔案存放路徑
        public static string strUploadPath = ConfigurationManager.AppSettings["UploadFolder"];
        public static string quotesFolder = "Quotes"; //廠商報價單路徑
        //Sample Code : It can get ADO.NET Dataset
        public DataSet ExecuteStoreQuery(string sql, CommandType commandType, Dictionary<string, Object> parameters)
        {
            var result = new DataSet();
            // creates a data access context (DbContext descendant)
            using (var context = new topmepEntities())
            {
                // creates a Command 
                var cmd = context.Database.Connection.CreateCommand();
                cmd.CommandType = commandType;
                cmd.CommandText = sql;

                // adds all parameters
                foreach (var pr in parameters)
                {
                    var p = cmd.CreateParameter();
                    p.ParameterName = pr.Key;
                    p.Value = pr.Value;
                    cmd.Parameters.Add(p);
                }

                try
                {
                    // executes
                    context.Database.Connection.Open();
                    var reader = cmd.ExecuteReader();

                    // loop through all resultsets (considering that it's possible to have more than one)
                    do
                    {
                        // loads the DataTable (schema will be fetch automatically)
                        var tb = new DataTable();
                        tb.Load(reader);
                        result.Tables.Add(tb);

                    } while (!reader.IsClosed);
                }
                finally
                {
                    // closes the connection
                    context.Database.Connection.Close();
                }
            }
            // returns the DataSet
            return result;
        }

    }
    /// <summary>
    /// System User service
    /// </summary>
    public class UserService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public SYS_USER loginUser;
        public List<SYS_FUNCTION> userPrivilege;

        /// <remarks>
        /// User Login by userid and passeword and get provilege informante
        /// </remarks>
        public SYS_USER Login(String userid, String passwd)
        {
            loginUser = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    loginUser = context.SYS_USER.SqlQuery("select u.* from SYS_USER u "
                        + "where u.USER_ID = @userid "
                        + "and u.PASSWORD = @passwd "
                       , new SqlParameter("userid", userid), new SqlParameter("passwd", passwd)).First();
                }
                catch (Exception e)
                {
                    logger.Error("login fail:" + e.StackTrace);
                    throw e;
                }
            }
            logger.Info("get user info=" + loginUser);
            if (null != loginUser)
            {
                getPrivilege(userid, passwd);
            }
            return loginUser;
        }

        private void getPrivilege(String userid, String passwd)
        {
            userPrivilege = null;
            using (var context = new topmepEntities())
            {
                userPrivilege = context.SYS_FUNCTION.SqlQuery("select f.* from SYS_FUNCTION f,SYS_PRIVILEGE p,SYS_ROLE r,SYS_USER u "
                    + "where u.ROLE_ID = r.ROLE_ID "
                    + "and r.ROLE_ID = p.ROLE_ID "
                    + "and p.FUNCTION_ID = f.FUNCTION_ID "
                    + "and u.USER_ID = @userid "
                    + "and u.PASSWORD = @passwd "
                    + "Order by MODULE_NAME Desc;", new SqlParameter("userid", userid), new SqlParameter("passwd", passwd)).ToList();
            }
            logger.Info("get functions count=" + userPrivilege.Count);
        }

    }
    #region 備標處理區段
    /***
     * 備標階段專案管理
     */
    public class TnderProject : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        TND_PROJECT project = null;
        string sno_key = "PROJ";
        public static string UploadFolder = null;
        public TnderProject()
        {
        }
        public int newProject(TND_PROJECT prj)
        {

            //1.建立專案基本資料
            logger.Info("create new project " + prj.ToString());
            project = prj;
            int i = 0;
            using (var context = new topmepEntities())
            {
                //2.取得專案編號
                SerialKeyService snoservice = new SerialKeyService();
                project.PROJECT_ID = snoservice.getSerialKey(sno_key);
                logger.Info("new projecgt object=" + project.ToString());
                context.TND_PROJECT.Add(project);
                //3.建立專案存取路徑
                string projectFolder = UploadFolder + "/" + project.PROJECT_ID;
                if (Directory.Exists(projectFolder))
                {
                    //資料夾存在
                    logger.Info("Directory Exist:" + projectFolder);
                }
                else
                {
                    //if directory not exist create it
                    Directory.CreateDirectory(projectFolder);
                }
                i = context.SaveChanges();
                logger.Debug("Add project=" + i);
            }
            return i;
        }
        public int updateProject(TND_PROJECT prj)
        {
            //1.建立專案基本資料
            project = prj;
            logger.Info("Update project " + project.ToString());
            int i = 0;
            using (var context = new topmepEntities())
            {
                context.Entry(project).State = EntityState.Modified;
                i = context.SaveChanges();
                logger.Debug("Update project=" + i);
            }
            return i;
        }
        #region 標單項目處理
        public int delAllItemByProject()
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all item by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM TND_PROJECT_ITEM WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", project.PROJECT_ID));
            }
            logger.Debug("delete item count=" + i);
            return i;
        }
        public int refreshProjectItem(List<TND_PROJECT_ITEM> prjItem)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshProjectItem = " + prjItem.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (TND_PROJECT_ITEM item in prjItem)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.TND_PROJECT_ITEM.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add project item count =" + i);
            return i;
        }
        #endregion
        //2.建立任務分配表
        TND_TASKASSIGN task = null;
        public void newTask(TND_TASKASSIGN task)
        {
            //1.建立專案基本資料
            logger.Info("create new task ");
            using (var context = new topmepEntities())
            {
                context.TND_TASKASSIGN.Add(task);
                int i = context.SaveChanges();
                logger.Debug("Add task=" + i);
                //if (i > 0) { status = true; };
            }
        }

        public TND_PROJECT getProjectById(string prjid)
        {
            using (var context = new topmepEntities())
            {
                project = context.TND_PROJECT.SqlQuery("select p.* from TND_PROJECT p "
                    + "where p.PROJECT_ID = @pid "
                   , new SqlParameter("pid", prjid)).First();
            }
            return project;
        }

        TND_PROJECT_FORM form = null;
        public string newForm(TND_PROJECT_FORM form, string[] lstItemId)
        {
            //1.建立詢價單價單樣本
            logger.Info("create new project form ");
            string sno_key = "PO";
            SerialKeyService snoservice = new SerialKeyService();
            form.FORM_ID = snoservice.getSerialKey(sno_key);
            logger.Info("new projecgt form =" + form.ToString());
            using (var context = new topmepEntities())
            {
                context.TND_PROJECT_FORM.Add(form);
                int i = context.SaveChanges();
                logger.Debug("Add form=" + i);
                logger.Info("project form id = " + form.FORM_ID);
                //if (i > 0) { status = true; };
                List<topmeperp.Models.TND_PROJECT_FORM_ITEM> lstItem = new List<TND_PROJECT_FORM_ITEM>();
                string ItemId = "";
                for (i = 0; i < lstItemId.Count(); i++)
                {
                    if (i < lstItemId.Count() - 1)
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                    }
                    else
                    {
                        ItemId = ItemId + "'" + lstItemId[i] + "'";
                    }
                }

                string sql = "INSERT INTO TND_PROJECT_FORM_ITEM (FORM_ID, PROJECT_ITEM_ID, TYPE_CODE, "
                    + "SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QTY, ITEM_UNIT_PRICE, ITEM_REMARK) "
                    + "SELECT '" + form.FORM_ID + "' as FORM_ID, PROJECT_ITEM_ID, TYPE_CODE_1 AS TYPE_CODE, "
                    + "TYPE_CODE_1 AS SUB_TYPE_CODE, ITEM_DESC, ITEM_UNIT, ITEM_QUANTITY, ITEM_UNIT_PRICE, ITEM_REMARK "
                    + "FROM TND_PROJECT_ITEM where PROJECT_ITEM_ID IN (" + ItemId + ")";
                logger.Info("sql =" + sql);
                var parameters = new List<SqlParameter>();
                i = context.Database.ExecuteSqlCommand(sql);
                return form.FORM_ID;
            }
        }
        public TND_PROJECT_FORM getProjectFormById(string prjid)
        {
            using (var context = new topmepEntities())
            {
                form = context.TND_PROJECT_FORM.SqlQuery("select pf.* from TND_PROJECT_FORM pf "
                    + "where pf.PROJECT_ID = @pid "
                   , new SqlParameter("pid", prjid)).First();
            }
            return form;
        }

        public List<TND_TASKASSIGN> getTaskById(string projectid)
        {
            List<TND_TASKASSIGN> lstTask = new List<TND_TASKASSIGN>();
            using (var context = new topmepEntities())
            {
                lstTask = context.TND_TASKASSIGN.SqlQuery("select t.* from TND_TASKASSIGN t "
                    + "where t.PROJECT_ID = @projectid "
                   , new SqlParameter("projectid", projectid)).ToList();
            }
            return lstTask;
        }
        public List<TND_PROJECT_FORM_ITEM> getFormItemById(string[] lstItemId)
        {
            List<TND_PROJECT_FORM_ITEM> lstFormItem = new List<TND_PROJECT_FORM_ITEM>();
            int i = 0;
            string ItemId = "";
            for (i = 0; i < lstItemId.Count(); i++)
            {
                if (i < lstItemId.Count() - 1)
                {
                    ItemId = ItemId + "'" + lstItemId[i] + "'" + ",";
                }
                else
                {
                    ItemId = ItemId + "'" + lstItemId[i] + "'";
                }
            }

            using (var context = new topmepEntities())
            {
                lstFormItem = context.TND_PROJECT_FORM_ITEM.SqlQuery("select f.* from  TND_PROJECT_FORM_ITEM f "
                    + "where f.PROJECT_ITEM_ID in (" + ItemId + ");"
                   , new SqlParameter("ItemId", ItemId)).ToList();
            }

            return lstFormItem;
        }
        //取得標單品項資料
        public List<TND_PROJECT_ITEM> getProjectItem(string projectid, string typeCode1, string typeCode2, string systemMain, string systemSub)
        {

            logger.Info("search projectitem by 九宮格 =" + typeCode1 + "search projectitem by 次九宮格 =" + typeCode2 + "search projectitem by 主系統 =" + systemMain + "search projectitem by 次系統 =" + systemSub);
            List<topmeperp.Models.TND_PROJECT_ITEM> lstItem = new List<TND_PROJECT_ITEM>();
            //處理SQL 預先填入專案代號,設定集合處理參數
            string sql = "SELECT * FROM TND_PROJECT_ITEM p WHERE p.PROJECT_ID =@projectid ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("projectid", projectid));
            //九宮格
            if (null != typeCode1 && typeCode1 != "")
            {
                sql = sql + "AND p.TYPE_CODE_1 LIKE @typeCode1 ";
                parameters.Add(new SqlParameter("typeCode1", "%" + typeCode1 + "%"));
            }
            //次九宮格
            if (null != typeCode2 && typeCode2 != "")
            {
                sql = sql + "AND p.TYPE_CODE_2 LIKE @typeCode2 ";
                parameters.Add(new SqlParameter("typeCode2", "%" + typeCode2 + "%"));
            }
            //主系統
            if (null != systemMain && systemMain != "")
            {
                sql = sql + "AND p.SYSTEM_MAIN LIKE @systemMain ";
                parameters.Add(new SqlParameter("systemMain", "%" + systemMain + "%"));
            }
            //次系統
            if (null != systemSub && systemSub != "")
            {
                sql = sql + "AND p.SYSTEM_SUB LIKE @systemSub ";
                parameters.Add(new SqlParameter("systemSub", "%" + systemSub + "%"));
            }

            using (var context = new topmepEntities())
            {
                lstItem = context.TND_PROJECT_ITEM.SqlQuery(sql, parameters.ToArray()).ToList();
            }
            logger.Info("get projectitem count=" + lstItem.Count);
            return lstItem;
        }

        #region 消防水圖算數量  
        //增加消防水圖算數量
        public int refreshMapFW(List<TND_MAP_FW> items)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshProjectItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (TND_MAP_FW item in items)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.TND_MAP_FW.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add TND_MAP_FW count =" + i);
            return i;
        }
        public int delMapFWByProject(string projectid)
        {
            logger.Info("remove all FW by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all TND_MAP_FW by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM  TND_MAP_FW WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete TND_MAP_FW count=" + i);
            return i;
        }
        #endregion
        #region 給排水圖算數量  
        //增加給排水圖算數量
        public int refreshMapPLU(List<TND_MAP_PLU> items)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshProjectItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (TND_MAP_PLU item in items)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.TND_MAP_PLU.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add TND_MAP_PLU count =" + i);
            return i;
        }
        public int delMapPLUByProject(string projectid)
        {
            logger.Info("remove all PLU by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all TND_MAP_PLU by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM  TND_MAP_PLU WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete TND_MAP_PLU count=" + i);
            return i;
        }
        #endregion
        #region 弱電管線圖算數量  
        //增加弱電管線圖算數量
        public int refreshMapLCP(List<TND_MAP_LCP> items)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshProjectItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (TND_MAP_LCP item in items)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.TND_MAP_LCP.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add TND_MAP_LCP count =" + i);
            return i;
        }
        public int delMapLCPByProject(string projectid)
        {
            logger.Info("remove all LCP by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all TND_MAP_LCP by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM  TND_MAP_LCP WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete TND_MAP_LCP count=" + i);
            return i;
        }
        #endregion
        #region 電氣管線圖算數量  
        //增加電氣管線圖算數量
        public int refreshMapPEP(List<TND_MAP_PEP> items)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshProjectItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (TND_MAP_PEP item in items)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.TND_MAP_PEP.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add TND_MAP_PEP count =" + i);
            return i;
        }
        public int delMapPEPByProject(string projectid)
        {
            logger.Info("remove all PEP by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all TND_MAP_PEP by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM  TND_MAP_PEP WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete TND_MAP_PEP count=" + i);
            return i;
        }
        #endregion
        #region 消防電圖算數量  
        //消防電圖算數量  
        public int refreshMapFP(List<TND_MAP_FP> items)
        {
            //1.檢查專案是否存在
            if (null == project) { throw new Exception("Project is not exist !!"); }
            int i = 0;
            logger.Info("refreshProjectItem = " + items.Count);
            //2.將Excel 資料寫入 
            using (var context = new topmepEntities())
            {
                foreach (TND_MAP_FP item in items)
                {
                    item.PROJECT_ID = project.PROJECT_ID;
                    context.TND_MAP_FP.Add(item);
                }
                i = context.SaveChanges();
            }
            logger.Info("add TND_MAP_FP count =" + i);
            return i;
        }
        public int delMapFPByProject(string projectid)
        {
            logger.Info("remove all FP by project ID=" + projectid);
            int i = 0;
            using (var context = new topmepEntities())
            {
                logger.Info("delete all TND_MAP_FP by proejct id=" + project.PROJECT_ID);
                i = context.Database.ExecuteSqlCommand("DELETE FROM  TND_MAP_FP WHERE PROJECT_ID=@projectid", new SqlParameter("@projectid", projectid));
            }
            logger.Debug("delete TND_MAP_FP count=" + i);
            return i;
        }
        #endregion
    }
  
    //詢價單資料提供作業
    public class InquiryFormService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public TND_PROJECT_FORM formInquiry = null;
        public List<TND_PROJECT_FORM_ITEM> formInquiryItem = null;
        //取得詢價單
        public void getInqueryForm(string formid)
        {
            logger.Info("get form : formid=" + formid);
            using (var context = new topmepEntities())
            {
                //取得詢價單檔頭資訊
                formInquiry = context.TND_PROJECT_FORM.SqlQuery("SELECT * FROM TND_PROJECT_FORM WHERE FORM_ID=@formid", new SqlParameter("formid", formid)).First();
                //取得詢價單明細
                formInquiryItem = context.TND_PROJECT_FORM_ITEM.SqlQuery("SELECT * FROM TND_PROJECT_FORM_ITEM WHERE FORM_ID=@formid", new SqlParameter("formid", formid)).ToList();
                logger.Debug("get form item count:" + formInquiryItem.Count);
            }
        }
        //取得專案詢價單樣板(供應商欄位為0)
        public List<TND_PROJECT_FORM> getFormTemplateByProject(string projectid)
        {
            logger.Info("get inquiry template by projectid=" + projectid);
            List <TND_PROJECT_FORM> lst = new List<TND_PROJECT_FORM>();
            using (var context = new topmepEntities())
            {
                //取得詢價單樣本資訊
                lst = context.TND_PROJECT_FORM.SqlQuery("SELECT * FROM TND_PROJECT_FORM WHERE SUPPLIER_ID IS NULL AND　PROJECT_ID=@projectid", 
                    new SqlParameter("projectid", projectid)).ToList();
            }
            return lst;
        }

    }
    #endregion
    #region 序號服務提供區段
    /*
     * 序號處理程序
     */
    public class SerialKeyService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public SerialKeyService()
        {
        }
        /*增加序號記錄欄位*/
        public bool addSerialKey(SYS_KEY_SERIAL serialKey)
        {
            bool status = false;
            using (var context = new topmepEntities())
            {
                context.SYS_KEY_SERIAL.Add(serialKey);
                //_db.AddToSYS_KEY_SERIAL(serialKey);
                int i = context.SaveChanges();
                logger.Debug("Add SerialKey : SerialKey=" + serialKey.KEY_ID + ",status=" + i);
                if (i > 0) { status = true; };
            }
            return status;
        }
        /* 依 KEY_ID 取得新序號(String)*/
        public string getSerialKey(string keyId)
        {
            logger.Debug("get new id by key");
            SYS_KEY_SERIAL SnKey = null;
            String sKey = null;
            using (var context = new topmepEntities())
            {
                //1.取得現有序號值
                string esql = @"SELECT * FROM SYS_KEY_SERIAL AS serialKey WHERE serialKey.KEY_ID=@keyId";
                SnKey = context.SYS_KEY_SERIAL.SqlQuery(esql, new SqlParameter("keyId", keyId)).First();

                logger.Debug("get new key :" + SnKey.KEY_ID + "=" + SnKey.KEY_NO);
                sKey = SnKey.KEY_NO.ToString().Trim();
                //2.將序號補0
                while ((sKey.Length + +SnKey.PREFIX.Length) < SnKey.KEY_LEN)
                {
                    sKey = "0" + sKey;
                }

                SnKey.KEY_NO = SnKey.KEY_NO + 1;
                int i = context.SaveChanges();
                logger.Info("Update SerialKey: Status =" + i);
                sKey = SnKey.PREFIX + sKey;
                logger.Info("New KEY :" + SnKey.KEY_ID + "=" + sKey);
            }
            return sKey;
        }
    }
    #endregion
    #region 系統管理區塊
    /*
     *使用者帳號管理 
     */
    public class UserManage : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //封裝供前端頁面調用
        public UserManageModels userManageModels = new UserManageModels();
        public string message = "";
        //取得所有角色資料
        public void getAllRole()
        {
            List<SYS_ROLE> lstRoles = null;
            using (var context = new topmepEntities())
            {
                try
                {
                    lstRoles = context.SYS_ROLE.SqlQuery("SELECT * FROM SYS_ROLE").ToList();
                    logger.Debug("get records=" + lstRoles.Count);
                    //將系統所有角色封裝供前端頁面調用
                    userManageModels.sysRole = lstRoles;
                }
                catch (Exception e)
                {
                    logger.Error("fail:" + e.StackTrace);
                }
            }
        }
        //新增帳號資料
        public int addNewUser(SYS_USER u)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.SYS_USER.AddOrUpdate(u);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("add new user id fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        //新增角色
       　public int addOrUpdateRole(SYS_ROLE role)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                try
                {
                    context.SYS_ROLE.AddOrUpdate(role);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("add new role fail:" + e.ToString());
                    logger.Error(e.StackTrace);
                    message = e.Message;
                }

            }
            return i;
        }
        //取得帳號資料
        public void getUserByCriteria(SYS_USER u, string roleid)
        {
            logger.Info("user=" + u.ToString() + ",roleId=" + roleid);
            List<SYS_USER> lstUser = new List<SYS_USER>();
            //處理SQL，預先埋入條件減少後續處理作業
            string sql = "SELECT USER_ID,USER_NAME,EMAIL,TEL,TEL_EXT,PASSWORD,FAX,MOBILE,CREATE_ID,CREATE_DATE,MODIFY_ID,MODIFY_DATE,"
                + "(SELECT ROLE_NAME FROM SYS_ROLE r WHERE r.ROLE_ID = u.ROLE_ID) ROLE_ID "
                + " FROM SYS_USER u WHERE 1=1 ";
            //定義參數: User ID , User Name, Tel,Roleid
            var parameters = new List<SqlParameter>();
            //處理帳號相關條件
            if (null != u)
            {
                //帳號
                logger.Debug("userID=" + u.USER_ID);
                if (null != u.USER_ID && u.USER_ID != "")
                {
                    sql = sql + "AND u.USER_ID= @userid ";
                    parameters.Add(new SqlParameter("userid", u.USER_ID));
                }
                //姓名
                logger.Debug("USER_NAME=" + u.USER_NAME);
                if (null != u.USER_NAME && u.USER_NAME != "")
                {
                    sql = sql + "AND u.USER_NAME LIKE  @username ";
                    parameters.Add(new SqlParameter("username", "%" + u.USER_NAME + "%"));
                }
                //電話
                logger.Debug("TEL=" + u.TEL);
                if (null != u.TEL && u.TEL != "")
                {
                    sql = sql + "AND u.TEL LIKE  @tel ";
                    parameters.Add(new SqlParameter("tel", "%" + u.TEL + "%"));
                }
            }
            //填入角色條件
            if (null != roleid && roleid != "")
            {
                logger.Debug("ROLE_ID=" + u.ROLE_ID);
                sql = sql + "AND u.ROLE_ID = @roleid ";
                parameters.Add(new SqlParameter("roleid", roleid));
            }
            //取得資料
            using (var context = new topmepEntities())
            {
                if (parameters.Count() == 0)
                {
                    logger.Debug(sql);
                    lstUser = context.SYS_USER.SqlQuery(sql).ToList();
                }
                else
                {
                    logger.Debug(sql);
                    lstUser = context.SYS_USER.SqlQuery(sql, parameters.ToArray()).ToList();
                }

                userManageModels.sysUsers = lstUser;
            }
        }
        public List<SYS_FUNCTION> getFunctions(string roleid)
        {
            List<SYS_FUNCTION> lstFunction = new List<SYS_FUNCTION>();
            logger.Info("roleid=" + roleid);
            //處理SQL
            string sql = "SELECT * FROM SYS_FUNCTION;";
            using (var context = new topmepEntities())
            {
                lstFunction = context.SYS_FUNCTION.SqlQuery(sql).ToList();
                logger.Debug("function count=" + lstFunction.Count);
            }
            return lstFunction;
        }
        //取得使用者資料
        public SYS_USER getUser(string userid)
        {
            logger.Debug("get user by id=" + userid);
            SYS_USER u = null;
            using (var context = new topmepEntities())
            {
                //設定此2參數，以便取消關聯物件，讓JSON 可以運作
                // Disable lazy loading
                context.Configuration.LazyLoadingEnabled = false;
                // Disable proxies
                context.Configuration.ProxyCreationEnabled = false;
                //設定SQL
                string esql = @"SELECT * FROM SYS_USER u WHERE u.USER_ID=@userid";
                try
                {
                    u = context.SYS_USER.SqlQuery(esql, new SqlParameter("userid", userid)).First();
                }
                catch (Exception e)
                {
                    logger.Error(e);
                }
            }
            return u;
        }
        public List<PrivilegeFunction> getPrivilege(string roleid)
        {
            List<PrivilegeFunction> lst = new List<PrivilegeFunction>();
            using (var context = new topmepEntities())
            {
                lst = context.Database.SqlQuery<PrivilegeFunction>("SELECT  f.FUNCTION_ID, f.FUNCTION_NAME, f.MODULE_NAME, f.FUNCTION_URI,"
                    + "f.CREATE_DATE, f.CREATE_ID, f.MODIFY_DATE, f.MODIFY_ID, r.ROLE_ID "
                    + "FROM  SYS_FUNCTION  f left outer join "
                    + "(SELECT ROLE_ID, FUNCTION_ID FROM SYS_PRIVILEGE p where p.ROLE_ID = @roleid) r "
                    + "on f.FUNCTION_ID = r.FUNCTION_ID;", new SqlParameter("roleid", roleid)).ToList();
            }
            logger.Info("get function count:" + lst.Count);
            return lst;
        }
    }
    #endregion
}
