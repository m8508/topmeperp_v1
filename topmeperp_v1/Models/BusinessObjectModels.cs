using System;
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

    public class MapInfoModels
    {
        //圖算消防電資料
        public IEnumerable<TND_MAP_FP> mapFP { get; set; }
        //圖算消防電資料
        public IEnumerable<TND_MAP_FW> mapFW { get; set; }
        //圖算給排水資料
        public IEnumerable<TND_MAP_PLU> mapPLU { get; set; }
        //圖算弱電管線資料
        public IEnumerable<TND_MAP_LCP> mapLCP { get; set; }
        //圖算電氣管線資料
        public IEnumerable<TND_MAP_PEP> mapPEP { get; set; }
        //圖算設備清單資料
        public IEnumerable<TND_MAP_DEVICE> mapDEVICE { get; set; }
    }
    public class InquiryFormModel
    {
        /// <summary>
        /// 詢價單樣本
        /// </summary>
        public IEnumerable<TND_PROJECT_FORM> tndTemplateProjectForm { get; set; }
        /// <summary>
        /// 供應商報價單
        /// </summary>
        public IEnumerable<TND_PROJECT_FORM> tndProjectFormFromSupplier { get; set; }
    }
}