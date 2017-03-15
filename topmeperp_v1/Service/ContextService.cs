using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
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
        public static string strUploadPath = "~/UploadFile";
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
            UploadFolder = ConfigurationManager.AppSettings["UploadFolder"];
            logger.Info("initial upload foler:" + UploadFolder);
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

        public TND_TASKASSIGN getTaskById(string taskid)
        {
            using (var context = new topmepEntities())
            {
                task = context.TND_TASKASSIGN.SqlQuery("select t.* from TND_TASKASSIGN t "
                    + "where t.PROJECT_ID = @taskid "
                   , new SqlParameter("taskid", taskid)).First();
            }
            return task;
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
    }
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
    /*
     *使用者帳號管理 
     */
    public class UserManage : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //封裝供前端頁面調用
        UserManageModels userManageModels = new UserManageModels();
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
                    context.SYS_USER.Add(u);
                    i = context.SaveChanges();
                }
                catch (Exception e)
                {
                    logger.Error("add new user id fail:" + e.Message);
                }

            }
            return i;
        }
        //取得帳號資料
        public void getUserByCriteria (SYS_USER u,string roleid)
        {
            logger.Info("Criteria= user=" + u.ToString() +",roleId="+roleid);
            List<SYS_USER> lstUser = new List<SYS_USER>();
            //處理SQL，預先埋入條件減少後續處理作業
            string sql = "SELECT * FROM SYS_USER u WHERE 1=1 ";
            //定義參數: User ID , User Name, Tel,Roleid
            var parameters = new List<SqlParameter>();
            //處理帳號相關條件
            if (null != u)
            {
                //帳號
                if (null != u.USER_ID && u.USER_ID != "")
                {
                    sql = sql + "AND u.USER_ID= @userid ";
                    parameters.Add(new SqlParameter("userid", u.USER_ID));
                }
                //姓名
                if (null != u.USER_NAME && u.USER_NAME != "")
                {
                    sql = sql + "AND u.USER_NAME LIKE  @username ";
                    parameters.Add(new SqlParameter("username", "%" + u.USER_NAME + "%"));
                }
                //電話
                if (null != u.TEL && u.TEL != "")
                {
                    sql = sql + "AND u.TEL LIKE  @tel ";
                    parameters.Add(new SqlParameter("tel", "%" + u.TEL + "%"));
                }
            }
            //填入角色條件
            if (null != roleid || roleid != "")
            {
                sql = sql + "AND u.ROLE_ID = @roleid ";
                parameters.Add(new SqlParameter("roleid", roleid));
            }
            //取得資料
            using (var context = new topmepEntities())
            {
                lstUser = context.SYS_USER.SqlQuery(sql, parameters.ToArray()).ToList();
                userManageModels.sysUsers = lstUser;
            } 
        }
    }
}
