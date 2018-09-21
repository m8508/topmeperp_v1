using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace topmeperp.Models
{
    //代付扣回資料
    public class Model4PaymentTransfer
    {
        public string EST_FORM_ID { get; set; }
        public string PROJECT_ID { get; set; }
        public string CONTRACT_ID { get; set; }
        public string HOLD4SUPPLIER { get; set; }
        public DateTime CREATE_DATE { get; set; }
        public Nullable<decimal> PAID_AMOUNT { get; set; }
        public string HOLD4REMARK { get; set; }
    }
}