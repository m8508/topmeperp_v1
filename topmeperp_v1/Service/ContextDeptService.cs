using log4net;
using Newtonsoft.Json;
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

    #region 部門資料管理管理區塊
    /*
     *部門資料管理管理 
     */
    public class DepartmentManage : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public TND_SUPPLIER supplier = null;
        //  public SUP_MATERIAL_RELATION typemain = null;
        public TND_SUP_CONTACT_INFO contact = null;
        public List<TND_SUP_CONTACT_INFO> contactList = null;
        string sno_key = "SUP";
    }
    #endregion


}