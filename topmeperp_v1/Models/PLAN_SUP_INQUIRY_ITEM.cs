//------------------------------------------------------------------------------
// <auto-generated>
//     這個程式碼是由範本產生。
//
//     對這個檔案進行手動變更可能導致您的應用程式產生未預期的行為。
//     如果重新產生程式碼，將會覆寫對這個檔案的手動變更。
// </auto-generated>
//------------------------------------------------------------------------------

namespace topmeperp.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class PLAN_SUP_INQUIRY_ITEM
    {
        public long INQUIRY_ITEM_ID { get; set; }
        public string INQUIRY_FORM_ID { get; set; }
        public string PLAN_ITEM_ID { get; set; }
        public string TYPE_CODE { get; set; }
        public string SUB_TYPE_CODE { get; set; }
        public string ITEM_ID { get; set; }
        public string ITEM_DESC { get; set; }
        public string ITEM_UNIT { get; set; }
        public Nullable<decimal> ITEM_QTY { get; set; }
        public Nullable<decimal> ITEM_UNIT_PRICE { get; set; }
        public Nullable<decimal> ITEM_QTY_ORG { get; set; }
        public Nullable<decimal> ITEM_UNITPRICE_ORG { get; set; }
        public string ITEM_REMARK { get; set; }
        public string MODIFY_ID { get; set; }
        public Nullable<System.DateTime> MODIFY_DATE { get; set; }
        public Nullable<decimal> ITEM_COUNTER_OFFER { get; set; }
    }
}
