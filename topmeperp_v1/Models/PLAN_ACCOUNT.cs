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
    
    public partial class PLAN_ACCOUNT
    {
        public long PLAN_ACCOUNT_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public string ACCOUNT_FORM_ID { get; set; }
        public Nullable<System.DateTime> PAYMENT_DATE { get; set; }
        public Nullable<decimal> AMOUNT_PAID { get; set; }
        public string ACCOUNT_TYPE { get; set; }
        public string ISDEBIT { get; set; }
        public Nullable<int> STATUS { get; set; }
        public string CREATE_ID { get; set; }
        public string MODIFY_ID { get; set; }
        public Nullable<System.DateTime> MODIFY_DATE { get; set; }
        public string CHECK_NO { get; set; }
        public Nullable<decimal> AMOUNT_PAYABLE { get; set; }
        public Nullable<decimal> AMOUNT { get; set; }
    }
}
