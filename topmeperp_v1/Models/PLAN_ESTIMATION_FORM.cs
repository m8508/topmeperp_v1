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
    
    public partial class PLAN_ESTIMATION_FORM
    {
        public string EST_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public string PLUS_TAX { get; set; }
        public Nullable<decimal> TAX_AMOUNT { get; set; }
        public Nullable<decimal> PAYMENT_TRANSFER { get; set; }
        public Nullable<decimal> OTHER_PAYMENT { get; set; }
        public Nullable<decimal> DEDUCTED_PAYMENT_TRANSFER { get; set; }
        public string REMARK { get; set; }
        public string CREATE_ID { get; set; }
        public Nullable<System.DateTime> CREATE_DATE { get; set; }
        public string SETTLEMENT { get; set; }
        public string TYPE { get; set; }
    }
}
