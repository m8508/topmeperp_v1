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
    public class PlanService : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


    }
    public class PurchaseService : TnderProject
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public List<string> getSystemMain(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SYSTEM_MAIN FROM PLAN_ITEM　WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get System Main Count=" + lst.Count);
            }
            return lst;
        }
        //取得次系統選單
        public List<string> getSystemSub(string projectid)
        {
            List<string> lst = new List<string>();
            using (var context = new topmepEntities())
            {
                //取得主系統選單
                lst = context.Database.SqlQuery<string>("SELECT DISTINCT SYSTEM_SUB FROM PLAN_ITEM WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                //lst = context.TND_PROJECT_ITEM.SqlQuery("SELECT DISTINCT SYSTEM_SUB FROM TND_PROJECT_ITEM　WHERE PROJECT_ID=@projectid;", new SqlParameter("projectid", projectid)).ToList();
                logger.Info("Get System Sub Count=" + lst.Count);
            }
            return lst;
        }
    }
}