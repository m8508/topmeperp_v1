using System;
using System.Collections.Generic;

namespace topmeperp.Models
{
    public class ProjectCompareData
    {
        public string SOURCE_PROJECT_ID { get; set; }
        public string SOURCE_SYSTEM_MAIN { get; set; }
        public string SOURCE_SYSTEM_SUB { get; set; }
        public string SOURCE_ITEM_ID { get; set; }
        public string SOURCE_ITEM_DESC { get; set; }
        public Nullable<decimal> SRC_UNIT_PRICE { get; set; }
        public string TARGET_PROJECT_ID { get; set; }
        public string TARGET_ITEM_ID { get; set; }
        public string TARGET_ITEM_DESC { get; set; }
        public string TARGET_SYSTEM_MAIN { get; set; }
        public string TARGET_SYSTEM_SUB { get; set; }
        public Nullable<long> EXCEL_ROW_ID { get; set; }
    }
}