using log4net;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Models;
using topmeperp.Service;
using System.IO;
namespace topmeperp.Controllers
{
    public class SupplierManageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(SupplierManageController));
       
        // GET: SupplierManage
        public ActionResult Index()
        {
            log.Info("index");
            return View();
        }
        //關鍵字尋找供應商
        public ActionResult Search()
        {
            List<SUPPLIER_FUNCTION> lstSupplier = SearchSupplierByName(Request["textSupplierName"],Request["textSuppyNote"], Request["textTypeMain"]);
            ViewBag.SearchResult = "共取得" + lstSupplier.Count + "筆資料";
            return View("Index", lstSupplier);
        }

        private List<SUPPLIER_FUNCTION> SearchSupplierByName(string suppliername,string supplyNote, string typeMain)
        {
            if (suppliername != null)
            {
                log.Info("search supplier by 名稱 =" + suppliername + ", by 九宮格 =" + typeMain + ", by 產品類別 =" + supplyNote);
                List <SUPPLIER_FUNCTION> lstSupplier = new List<SUPPLIER_FUNCTION>();
                string sql = "SELECT A.*, sc.CONTACT_ID, sc.SUPPLIER_MATERIAL_ID, sc.CONTACT_NAME, sc.CONTACT_TEL, sc.CONTACT_FAX, sc.CONTACT_EMAIL, " +
                    "sc.REMARK FROM (select s.*, sm.RELATION_ID, sm.TYPE_MAIN, sm.SUPPLY_NOTE, ISNULL(sm.STOP_DELIVERY, '供貨') AS STOP_DELIVERY " +
                    "from SUPPLIER s LEFT JOIN SUP_MATERIAL_RELATION sm ON s.SUPPLIER_ID = sm.SUPPLIER_ID) A LEFT JOIN SUP_CONTACT_INFO sc " +
                    "ON A.SUPPLIER_ID + A.TYPE_MAIN = sc.SUPPLIER_MATERIAL_ID WHERE A.SUPPLIER_ID IS NOT NULL  ";

                var parameters = new List<SqlParameter>();
                //九宮格條件
                if (null != typeMain && "" != typeMain)
                {
                    //sql = sql + " AND sm.TYPE_MAIN ='" + typeMain + "'";
                    sql = sql + " AND A.TYPE_MAIN=@typeMain";
                    parameters.Add(new SqlParameter("typeMain", typeMain));
                }
                //供應商名稱條件
                if (null != suppliername && "" != suppliername)
                {
                    //sql = sql + " AND s.COMPANY_NAME LIKE '%' + @suppliername + '%' ;
                    sql = sql + " AND A.COMPANY_NAME LIKE @suppliername";
                    parameters.Add(new SqlParameter("suppliername", '%' + @suppliername + '%'));
                }
                //產品類別條件
                if (null != supplyNote && "" != supplyNote)
                {
                    //sql = sql + " AND sm.SUPPLY_NOTE LIKE '%' + @supplyNote + '%' ;
                    sql = sql + " AND A.SUPPLY_NOTE LIKE @supplyNote";
                    parameters.Add(new SqlParameter("supplyNote", '%' + @supplyNote + '%'));
                }
                sql = sql + " ORDER BY A.TYPE_MAIN ";
                log.Info("sql=" + sql);
                using (var context = new topmepEntities())
                {
                    lstSupplier = context.Database.SqlQuery<SUPPLIER_FUNCTION>(sql, parameters.ToArray()).ToList();
                   
                }
                log.Info("get supplier count=" + lstSupplier.Count);
                return lstSupplier;
            }
            else
            {
                return null;
            }
        }
        //新增供應商資料
        //POST:Create
        public ActionResult Create(string id)
        {
            log.Info("get supplier for update:supplier_id = " + id) ;
            SUPPLIER s = null;
            if (null != id)
            {
                SupplierManage supplierService = new SupplierManage();
                s = supplierService.getSupplierById(id);
            }
            ViewBag.supplier = s.SUPPLIER_ID;
            return View(s);
        }
        [HttpPost]
        public ActionResult Create(SUPPLIER sup, SUP_MATERIAL_RELATION sm, SUP_CONTACT_INFO sc, FormCollection form)
        {
            log.Info("create supplier process! supplier =" + sup.ToString());
            log.Info("form:" + form.Count);
            SupplierManage supplierService = new SupplierManage();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //1.更新或新增專案基本資料
            if (sup.SUPPLIER_ID == "" || sup.SUPPLIER_ID == null)
            {
                //新增供應商九宮格
                sm.TYPE_MAIN = form["type_main"];
                sm.SUPPLY_NOTE = form["supply_note"];
                try
                {
                    sm.TYPE_SUB = int.Parse(form["type_sub"]);
                }
                catch (Exception ex)
                {
                    log.Error(sup.SUPPLIER_ID + " not type sub:" + ex.Message);
                }
                //新增供應商聯絡人
                sc.CONTACT_NAME = form["contact_name"];
                sc.CONTACT_TEL = form["contact_tel"];
                sc.CONTACT_FAX = form["contact_fax"];
                sc.CONTACT_MOBIL = form["contact_mobil"];
                sc.CONTACT_EMAIL = form["contact_email"];
                sc.REMARK = form["remark"];
                //新增供應商
                supplierService.newSupplier(sup, sm, sc);
                message = "新增一供應商 : 供應商編號" + sup.SUPPLIER_ID + "，其供應商名稱為" + sup.COMPANY_NAME + "，其九宮格的供應商編號為" + sm.SUPPLIER_ID + "，其九宮格與供應商編號為" + sc.SUPPLIER_MATERIAL_ID;
            }
            else
            {
                //修改供應商基本資料
                supplierService.updateSupplier(sup);
                message = "供應商基本資料修改成功 : 其供應商編號" + sup.SUPPLIER_ID + "，供應商名稱為" + sup.COMPANY_NAME;
            }
            ViewBag.supplier = sup.SUPPLIER_ID;
            ViewBag.result = message;
            return View(sup);
        }
        //編輯供應商九宮格相關資料
        public ActionResult EditForTypeMain(string id)
        {
            log.Info("get typeMain for update:relation_id=" + id);
            SUP_MATERIAL_RELATION sm = null;
            if (null != id)
            {
                SupplierManage supplierService = new SupplierManage();
                sm = supplierService.getTypeMainById(id);
            }
            ViewBag.stopdelivery = sm.STOP_DELIVERY;
            return View(sm);
        }
        [HttpPost]
        public ActionResult EditForTypeMain(SUP_MATERIAL_RELATION sm)
        {
            log.Info("update typeMain :supplier_id =" + sm.SUPPLIER_ID);
            string message = "";
            if (null != Request["stopDelivery"])
            {
                sm.STOP_DELIVERY = Request["stopDelivery"];
            }
            SupplierManage supplierService = new SupplierManage();
            supplierService.updateTypeMain(sm);
            message = "九宮格資料修改成功，SUPPLIER_ID :" + sm.SUPPLIER_ID;
            ViewBag.result = message;
            return View(sm);
        }
        /// <summary>
        /// 取得九宮格詳細資料
        /// </summary>
        /// <param name="relationid"></param>
        /// <returns></returns>
        public string getTypeMain(string relationid)
        {
            SupplierManage supplierService = new SupplierManage();
            log.Info("get type main info by supplier id=" + relationid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(supplierService.getTypeMainById(relationid));
            log.Info("supplier's type main  info=" + itemJson);
            return itemJson;
        }
        
        //新增供應商九宮格
        public String addTypeMain(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "新增成功!!";

            SUP_MATERIAL_RELATION item = new SUP_MATERIAL_RELATION();
            item.SUPPLIER_ID = form["supplier_id"];
            item.TYPE_MAIN = form["type_main"];
            item.SUPPLY_NOTE = form["supply_note"];
            try
            {
                item.TYPE_SUB = int.Parse(form["type_sub"]);
            }
            catch (Exception ex)
            {
                log.Error(item.SUPPLIER_ID + " not type sub:" + ex.Message);
            }
            SupplierManage supplierService = new SupplierManage();
            int i = supplierService.refreshTypeMain(item);
            if (i == 0) { msg = supplierService.message; }
            return msg;
        }
        //編輯供應商聯絡人相關資料
        public ActionResult EditForContact(string id)
        {
            log.Info("get contact for update:contact_id=" + id);
            SUP_CONTACT_INFO sc = null;
            if (null != id)
            {
                SupplierManage supplierService = new SupplierManage();
                sc = supplierService.getContactById(id);
            }
            return View(sc);
        }
        [HttpPost]
        public ActionResult EditForContact(SUP_CONTACT_INFO sc)
        {
            log.Info("update contact :supplier_material_id =" + sc.SUPPLIER_MATERIAL_ID);
            string message = "";
            SupplierManage supplierService = new SupplierManage();
            supplierService.updateContact(sc);
            message = "聯絡人資料修改成功，SUPPLIER_MATERIAL_ID :" + sc.SUPPLIER_MATERIAL_ID;
            ViewBag.result = message;
            return View(sc);
        }
        /// <summary>
        /// 取得聯絡人詳細資料
        /// </summary>
        /// <param name="contactid"></param>
        /// <returns></returns>
        public string getContact(string contactid)
        {
            SupplierManage supplierService = new SupplierManage();
            log.Info("get contact info by contact id=" + contactid);
            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            string itemJson = objSerializer.Serialize(supplierService.getContactById(contactid));
            log.Info("supplier's type main  info=" + itemJson);
            return itemJson;
        }
        //新增供應聯絡人
        public String addContact(FormCollection form)
        {
            log.Info("form:" + form.Count);
            string msg = "新增成功!!";

            SUP_CONTACT_INFO item = new SUP_CONTACT_INFO();
            item.SUPPLIER_MATERIAL_ID = form["supplier_material_id"];
            item.CONTACT_NAME = form["contact_name"];
            item.CONTACT_TEL = form["contact_tel"];
            item.CONTACT_FAX = form["contact_fax"];
            item.CONTACT_MOBIL = form["contact_mobil"];
            item.CONTACT_EMAIL = form["contact_email"];
            item.REMARK = form["remark"];
            SupplierManage supplierService = new SupplierManage();
            int i = supplierService.refreshContact(item);
            if (i == 0) { msg = supplierService.message; }
            return msg;
        }

}
}
