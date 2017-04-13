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
    #region 系統管理相關
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
    #endregion
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
    /// <summary>
    /// 提供詢價單、報價單所需資料結構
    /// </summary>
    public class InquiryFormDetail
    {
        public TND_PROJECT prj { get; set; }
        public TND_PROJECT_FORM prjForm { get; set; }
        public IEnumerable<TND_PROJECT_FORM_ITEM> prjFormItem { get; set; }
    }
    public class COMPARASION_DATA : TND_PROJECT_ITEM
    {
        //報價單編號
        public string FORM_ID { get; set; }
        //供應商報價
        public Nullable<decimal> QUOTATION_PRICE { get; set; }
        //供應商提供數量(理論上同詢價單價格)
        public Nullable<decimal> OFFER_QTY { get; set; }
        //供應商名稱(尚未由供應商主檔取得)，需另外調整
        public string SUPPLIER_NAME { get; set; }
    }
    public class SupplierManageModels
    {
        //供應商資料
        public IEnumerable<TND_SUPPLIER> suppliers { get; set; }
    }
    public class DirectCost
    {
        /// 九宮格編碼長度 2 
        public string MAINCODE { get; set; }
        /// 九宮格名稱
        public string MAINCODE_DESC { get; set; }
        //完整地次九宮格編碼
        public Nullable<int> T_SUB_CODE { get; set; }
        //次九宮格編碼
        public string SUB_CODE { get; set; }
        //次九宮格名稱
        public string SUB_DESC { get; set; }
        public Nullable<decimal> MATERIAL_COST { get; set; }
        public Nullable<decimal> MAN_DAY { get; set; }
        public Nullable<int> ITEM_COUNT { get; set; }
    }
    public class SystemCost
    {
        /// 九宮格編碼長度 2 
        public string SYSTEM_MAIN { get; set; }
        /// 九宮格名稱
        public string SYSTEM_SUB { get; set; }
        public Nullable<decimal> MATERIAL_COST { get; set; }
        public Nullable<decimal> MAN_DAY { get; set; }
        public Nullable<int> ITEM_COUNT { get; set; }
    }
}