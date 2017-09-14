using System;
using System.Collections.Generic;

namespace topmeperp.Models
{
    //採發階段之發包標的
    public class PURCHASE_ORDER: PLAN_SUP_INQUIRY
    {
        public Nullable<decimal> BudgetAmount { get; set; }
        public Nullable<int> CountPO { get; set; }
    }
}