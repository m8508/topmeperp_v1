﻿//------------------------------------------------------------------------------
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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class topmepEntities : DbContext
    {
        public topmepEntities()
            : base("name=topmepEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<REF_TYPE_MAIN> REF_TYPE_MAIN { get; set; }
        public virtual DbSet<REF_TYPE_SUB> REF_TYPE_SUB { get; set; }
        public virtual DbSet<SYS_KEY_SERIAL> SYS_KEY_SERIAL { get; set; }
        public virtual DbSet<SYS_PRIVILEGE> SYS_PRIVILEGE { get; set; }
        public virtual DbSet<SYS_ROLE> SYS_ROLE { get; set; }
        public virtual DbSet<SYS_USER> SYS_USER { get; set; }
        public virtual DbSet<TND_FILE> TND_FILE { get; set; }
        public virtual DbSet<TND_MAP_FP> TND_MAP_FP { get; set; }
        public virtual DbSet<TND_MAP_FW> TND_MAP_FW { get; set; }
        public virtual DbSet<TND_MAP_PEP> TND_MAP_PEP { get; set; }
        public virtual DbSet<TND_WAGE> TND_WAGE { get; set; }
        public virtual DbSet<TND_PROJECT_FORM_ITEM> TND_PROJECT_FORM_ITEM { get; set; }
        public virtual DbSet<TND_MAP_LCP> TND_MAP_LCP { get; set; }
        public virtual DbSet<TND_MAP_PLU> TND_MAP_PLU { get; set; }
        public virtual DbSet<TND_MAP_DEVICE> TND_MAP_DEVICE { get; set; }
        public virtual DbSet<TND_TASKASSIGN> TND_TASKASSIGN { get; set; }
        public virtual DbSet<TND_PROJECT_ITEM> TND_PROJECT_ITEM { get; set; }
        public virtual DbSet<TND_PROJECT_FORM> TND_PROJECT_FORM { get; set; }
        public virtual DbSet<PLAN_BUDGET> PLAN_BUDGET { get; set; }
        public virtual DbSet<PLAN_PAYMENT_TERMS> PLAN_PAYMENT_TERMS { get; set; }
        public virtual DbSet<PLAN_ITEM> PLAN_ITEM { get; set; }
        public virtual DbSet<TND_PROJECT> TND_PROJECT { get; set; }
        public virtual DbSet<PLAN_TASK> PLAN_TASK { get; set; }
        public virtual DbSet<TND_SUPPLIER> TND_SUPPLIER { get; set; }
        public virtual DbSet<TND_SUP_CONTACT_INFO> TND_SUP_CONTACT_INFO { get; set; }
        public virtual DbSet<PLAN_TASK2MAPITEM> PLAN_TASK2MAPITEM { get; set; }
        public virtual DbSet<TND_SUP_MATERIAL_RELATION> TND_SUP_MATERIAL_RELATION { get; set; }
        public virtual DbSet<PLAN_PURCHASE_REQUISITION_ITEM> PLAN_PURCHASE_REQUISITION_ITEM { get; set; }
        public virtual DbSet<SYS_FUNCTION> SYS_FUNCTION { get; set; }
        public virtual DbSet<PLAN_PURCHASE_REQUISITION> PLAN_PURCHASE_REQUISITION { get; set; }
        public virtual DbSet<PLAN_ITEM_DELIVERY> PLAN_ITEM_DELIVERY { get; set; }
        public virtual DbSet<PLAN_DALIY_REPORT> PLAN_DALIY_REPORT { get; set; }
        public virtual DbSet<PLAN_DR_ITEM> PLAN_DR_ITEM { get; set; }
        public virtual DbSet<PLAN_DR_NOTE> PLAN_DR_NOTE { get; set; }
        public virtual DbSet<PLAN_DR_TASK> PLAN_DR_TASK { get; set; }
        public virtual DbSet<PLAN_DR_WORKER> PLAN_DR_WORKER { get; set; }
        public virtual DbSet<SYS_PARA> SYS_PARA { get; set; }
        public virtual DbSet<PLAN_SUP_INQUIRY> PLAN_SUP_INQUIRY { get; set; }
        public virtual DbSet<PLAN_SUP_INQUIRY_ITEM> PLAN_SUP_INQUIRY_ITEM { get; set; }
        public virtual DbSet<PLAN_ESTIMATION_FORM> PLAN_ESTIMATION_FORM { get; set; }
        public virtual DbSet<PLAN_ESTIMATION_ITEM> PLAN_ESTIMATION_ITEM { get; set; }
    }
}
