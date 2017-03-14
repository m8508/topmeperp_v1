using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
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
        public List<TND_PROJECT_ITEM> getProjectItem(string projectid,string projectitem1, string projectitem2, string projectitem3, string projectitem4)
        {

            logger.Info("search projectitem by 九宮格 =" + projectitem1 + "search projectitem by 次九宮格 =" + projectitem2 + "search projectitem by 主系統 =" + projectitem3 + "search projectitem by 次系統 =" + projectitem4);
            List<topmeperp.Models.TND_PROJECT_ITEM> lstItem = new List<TND_PROJECT_ITEM>();
            using (var context = new topmepEntities())
            {
                lstItem = context.TND_PROJECT_ITEM.SqlQuery("select * from TND_PROJECT_ITEM p "
                    + "where p.TYPE_CODE_1 Like @typecode1 "
                    + "or p.TYPE_CODE_2 Like @typecode2 "
                    + "or p.SYSTEM_MAIN Like @systemMain "
                    + "or p.SYSTEM_SUB Like  @systemSub",
                     new SqlParameter("typecode1", "%" + projectitem1 + "%"),
                     new SqlParameter("typecode2", "%" + projectitem2 + "%"),
                     new SqlParameter("systemMain", "%" + projectitem3 + "%"),
                     new SqlParameter("systemSub", "%" + projectitem4 + "%")).ToList();
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
}
