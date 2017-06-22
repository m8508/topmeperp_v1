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
            List<SUPPLIER_FUNCTION> lstSupplier = SearchSupplierByName(Request["textSupplierName"], Request["textSuppyNote"], Request["textTypeMain"]);
            ViewBag.SearchResult = "共取得" + lstSupplier.Count + "筆資料";
            return View("Index", lstSupplier);
        }

        private List<SUPPLIER_FUNCTION> SearchSupplierByName(string suppliername, string supplyNote, string typeMain)
        {
            if (suppliername != null)
            {
                log.Info("search supplier by 名稱 =" + suppliername + ", by 九宮格 =" + typeMain + ", by 產品類別 =" + supplyNote);
                List<SUPPLIER_FUNCTION> lstSupplier = new List<SUPPLIER_FUNCTION>();
                string sql = "SELECT s.SUPPLIER_ID, s.COMPANY_NAME, s.COMPANY_ID, s.REGISTER_ADDRESS, s.CONTACT_ADDRESS, s.TYPE_MAIN, s.TYPE_SUB, " +
                    "s.SUPPLY_NOTE, sc.CONTACT_ID, sc.SUPPLIER_MATERIAL_ID, sc.CONTACT_NAME, sc.CONTACT_TEL, sc.CONTACT_FAX, sc.CONTACT_EMAIL, " +
                    "sc.REMARK FROM TND_SUPPLIER s RIGHT OUTER  JOIN TND_SUP_CONTACT_INFO sc " +
                    "ON s.SUPPLIER_ID = sc.SUPPLIER_MATERIAL_ID WHERE s.SUPPLIER_ID IS NOT NULL   ";

                var parameters = new List<SqlParameter>();
                //九宮格條件
                if (null != typeMain && "" != typeMain)
                {
                    //sql = sql + " AND sm.TYPE_MAIN ='" + typeMain + "'";
                    sql = sql + " AND s.TYPE_MAIN=@typeMain";
                    parameters.Add(new SqlParameter("typeMain", typeMain));
                }
                //供應商名稱條件
                if (null != suppliername && "" != suppliername)
                {
                    //sql = sql + " AND s.COMPANY_NAME LIKE '%' + @suppliername + '%' ;
                    sql = sql + " AND s.COMPANY_NAME LIKE @suppliername";
                    parameters.Add(new SqlParameter("suppliername", '%' + @suppliername + '%'));
                }
                //產品類別條件
                if (null != supplyNote && "" != supplyNote)
                {
                    //sql = sql + " AND sm.SUPPLY_NOTE LIKE '%' + @supplyNote + '%' ;
                    sql = sql + " AND s.SUPPLY_NOTE LIKE @supplyNote";
                    parameters.Add(new SqlParameter("supplyNote", '%' + @supplyNote + '%'));
                }
                sql = sql + " ORDER BY s.TYPE_MAIN ";
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
            log.Info("get supplier for update:supplier_id = " + id);
            TND_SUPPLIER s = null;
            if (null != id)
            {
                SupplierManage supplierService = new SupplierManage();
                s = supplierService.getSupplierById(id);
                ViewBag.supplier = s.SUPPLIER_ID;
            }
            return View(s);
        }
        [HttpPost]
        public ActionResult Create(TND_SUPPLIER sup, TND_SUP_CONTACT_INFO sc, FormCollection form)
        {
            log.Info("create supplier process! supplier =" + sup.ToString());
            log.Info("form:" + form.Count);
            SupplierManage supplierService = new SupplierManage();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //1.更新或新增專案基本資料
            if (sup.SUPPLIER_ID == "" || sup.SUPPLIER_ID == null)
            {
                //新增供應商聯絡人
                sc.CONTACT_NAME = form["contact_name"];
                sc.CONTACT_TEL = form["contact_tel"];
                sc.CONTACT_FAX = form["contact_fax"];
                sc.CONTACT_MOBIL = form["contact_mobil"];
                sc.CONTACT_EMAIL = form["contact_email"];
                sc.REMARK = form["remark"];
                //新增供應商
                supplierService.newSupplier(sup, sc);
                message = "新增一供應商 : 供應商編號" + sup.SUPPLIER_ID + "，其供應商名稱為" + sup.COMPANY_NAME + "，其九宮格與供應商編號為" + sc.SUPPLIER_MATERIAL_ID;
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

        //編輯供應商聯絡人相關資料
        public ActionResult EditForContact(string id)
        {
            log.Info("get contact for update:contact_id=" + id);
            TND_SUP_CONTACT_INFO sc = null;
            if (null != id)
            {
                SupplierManage supplierService = new SupplierManage();
                sc = supplierService.getContactById(id);
            }
            return View(sc);
        }
        [HttpPost]
        public ActionResult EditForContact(TND_SUP_CONTACT_INFO sc)
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

            TND_SUP_CONTACT_INFO item = new TND_SUP_CONTACT_INFO();
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
