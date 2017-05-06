using log4net;
using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.Linq;
using topmeperp.Models;

namespace topmeperp.Service
{
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
        public int updatePrivilege(string roleid, string[] functions)
        {
            int i = 0;
            using (var context = new topmepEntities())
            {
                //1.移除該角色所有權限
                string sql = "DELETE FROM SYS_PRIVILEGE WHERE ROLE_ID=@roleid;";
                i = context.Database.ExecuteSqlCommand(sql, new SqlParameter("roleid", roleid));
                logger.Info("Remove privilege count=" + i);
                //2.逐一加入授權資料
                for (int j = 0; j < functions.Length; j++)
                {
                    SYS_PRIVILEGE p = new SYS_PRIVILEGE();
                    p.PRIVILEGE_ID = roleid + "-" + functions[j];
                    p.ROLE_ID = roleid;
                    p.FUNCTION_ID = functions[j];
                    context.SYS_PRIVILEGE.Add(p);
                }
                i = context.SaveChanges();
                logger.Info("create privlilege data count:" + i);
            }
            return i;
        }
    }
    #endregion
    #region 九宮格次九宮格管理區塊
    public class TypeManageService
    {
        TyepManageModel typeManageModel = null;
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public List<REF_TYPE_MAIN> getMainType()
        {
            List<REF_TYPE_MAIN> lst = new List<REF_TYPE_MAIN>();
            using (var context = new topmepEntities())
            {
                lst = context.REF_TYPE_MAIN.SqlQuery("SELECT * FROM REF_TYPE_MAIN ORDER BY TYPE_CODE_1 ,TYPE_CODE_2;").ToList();
                logger.Debug("get Main Type Count" + lst.Count);
            }
            return lst;
        }
        public List<REF_TYPE_SUB> getSubType(string typecodeid)
        {
            List<REF_TYPE_SUB> lst = new List<REF_TYPE_SUB>();
            using (var context = new topmepEntities())
            {
                lst = context.REF_TYPE_SUB.SqlQuery("SELECT * FROM REF_TYPE_SUB WHERE TYPE_CODE_ID=@typecodeid ORDER BY SUB_TYPE_CODE;", new SqlParameter("typecodeid", typecodeid)).ToList();
                logger.Debug("get sub Type Count" + lst.Count);
            }
            return lst;
        }
        public TyepManageModel getTypeManageModel(string typecode1,string typecode2)
        {
            TyepManageModel typeManageModel = new TyepManageModel();
            //List<REF_TYPE_MAIN> lstMainType = new List<REF_TYPE_MAIN>();
            using (var context = new topmepEntities())
            {
                var parameters = new List<SqlParameter>();
                string sql = "SELECT * FROM REF_TYPE_MAIN WHERE TYPE_CODE_1=@code1 AND TYPE_CODE_2=@code2 ";
                logger.Debug(sql);
                parameters.Add(new SqlParameter("code1", typecode1));
                parameters.Add(new SqlParameter("code2", typecode2));
                //lstMainType = context.REF_TYPE_MAIN.SqlQuery(sql, parameters.ToArray()).ToList();
                typeManageModel.MainType = context.REF_TYPE_MAIN.SqlQuery(sql, parameters.ToArray()).First();
            }
            typeManageModel.SubTypes = getSubType(typeManageModel.MainType.TYPE_CODE_1 + typeManageModel.MainType.TYPE_CODE_2);
            logger.Debug("get Main Type Count" + typeManageModel.MainType.TYPE_DESC + ",sub type count" + typeManageModel.SubTypes.Count());
            return typeManageModel;
        }
        public void updateTypeManageModel(REF_TYPE_MAIN mainType, List<REF_TYPE_SUB> lstSubType)
        {
            using (var context = new topmepEntities())
            {
                //修改九宮格內容
                context.REF_TYPE_MAIN.AddOrUpdate(mainType);
                int i = 0;
                i= context.SaveChanges();
                logger.Debug("Modify MainType :" + i);
                string sql = "DELETE FROM REF_TYPE_SUB WHERE TYPE_CODE_ID=@typecodeid";
                logger.Debug("Remove SubType=" + sql + ",TypeCode="+ mainType.TYPE_CODE_1 + mainType.TYPE_CODE_2);
                context.Database.ExecuteSqlCommand(sql, new SqlParameter("typecodeid", mainType.TYPE_CODE_1+ mainType.TYPE_CODE_2));
                
                foreach (REF_TYPE_SUB subType in lstSubType)
                {
                    logger.Debug("subTypeId=" + subType.SUB_TYPE_ID + ",mainTypeId=" + subType.TYPE_CODE_ID + ",subTypeCode=" + subType.SUB_TYPE_CODE + ",subTypeDesc=" + subType.TYPE_DESC);
                    subType.CREATE_DATE = System.DateTime.Now;
                    context.REF_TYPE_SUB.Add(subType);
                }
                i = context.SaveChanges();
                logger.Debug("Modify MainSub :" + i);
            }
        }
    }
    #endregion
}