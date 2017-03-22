﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    //定義針對特定View 所需的資料集合
    public class BusinessObjectModels
    {
    }
    //**備標階段標書基本資料(不包含圖算數量)
    public class TndProjectModels
    {
        //標單專案檔
        public TND_PROJECT tndProject { get; set; }
        //標單上品項明細資料
        public IEnumerable<TND_PROJECT_ITEM> tndProjectItem { get; set; }
        //專案任務分工
        public IEnumerable<TND_TASKASSIGN> tndTaskAsign { get; set; }
        //專案相關檔案
        public IEnumerable<TND_FILE> tndFile { get; set; }
    }
    public class UserManageModels
    {
        //帳號資料
        public IEnumerable<SYS_USER> sysUsers { get; set; }
        //角色資料
        public IEnumerable<SYS_ROLE> sysRole { get; set; }
    }
    public class PrivilegeFunction : SYS_FUNCTION
    {
        public string ROLE_ID { get; set; }
    }

    public class TndFormModels
    {
        //詢價單表頭資料
        public IEnumerable<TND_PROJECT_FORM> tndProjectForm { get; set; }
        //詢價單明細資料
        public IEnumerable<TND_PROJECT_FORM_ITEM> tndProjectFormItem { get; set; }
    }
}