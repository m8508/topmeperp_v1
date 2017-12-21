﻿using System;
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
    /***
     * 施工日報完整物件
     */
    public class DailyReport
    {
        public TND_PROJECT project { get; set; }
        public PLAN_DALIY_REPORT dailyRpt { get; set; }
        //for display
        public List<DailyReportItem> lstDailyRptItem4Show { get; set; }

        public List<DailyReportRecord4Worker> lstDailyRptWokerType4Show { get; set; }
        public List<DailyReportRecord4Worker> lstDailyRptMachine4Show { get; set; }
        //phyical data
        public List<PLAN_DR_TASK> lstRptTask { get; set; }
        public List<PLAN_DR_ITEM> lstRptItem { get; set; }
        public List<PLAN_DR_WORKER> lstRptWorkerAndMachine { get; set; }
        public List<PLAN_DR_NOTE> lstRptNote { get; set; }
    }
    /// <summary>
    /// 施工日報人工與機具資料物件
    /// </summary>
    /// <param name="projectid"></param>
    /// <param name="prjuid"></param>
    /// <returns></returns>
    public class DailyReportRecord4Worker
    {
        public string FUNCTION_ID { get; set; }
        public string KEY_FIELD { get; set; }
        public string VALUE_FIELD { get; set; }
        public Nullable<decimal> LAST_QTY { get; set; }
        public Nullable<decimal> WORKER_QTY { get; set; }
        public string REMARK { get; set; }
        public string REPORT_ID { get; set; }
    }
    /// <summary>
    /// 成本預算管制表相關物件(缺追加減、成本異動資料
    /// </summary>
    public class CostControlInfo
    {
        public TND_PROJECT Project;
        public PlanRevenue Revenue;
        public  List<purchasesummary> lstDirectCostItem;
        public List<PLAN_INDIRECT_COST > lstIndirectCostItem;
    }
}