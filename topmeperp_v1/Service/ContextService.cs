using log4net;
using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Objects;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class ContextService
    {
        public topmepEntities db;// = new topmepEntities();

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
        public void newProject(TND_PROJECT prj)
        {
            //1.建立專案基本資料
            logger.Info("create new project ");
            using (var context = new topmepEntities())
            {
                context.TND_PROJECT.Add(prj);
                int i = context.SaveChanges();
                logger.Debug("Add project=" + i);
                //if (i > 0) { status = true; };
            }
            //2.建立任務分配表
        }
        public void newTask(TND_TASKASSIGN task)
        {
            logger.Info("assign task");
            using (var taskassign = new topmepEntities())
            {
                taskassign.TND_TASKASSIGN.Add(task);
                int j = taskassign.SaveChanges();
                logger.Debug("Add taskitem=" + j);
                //if (j > 0) { status = true; };
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
        public void impProjectItem()
        {
            ///1.匯入Excel 內容
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
                string esql = @"SELECT VALUE serialKey FROM TOPMEPEntities.SYS_KEY_SERIAL AS serialKey WHERE serialKey.KEY_ID=@keyId";
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
