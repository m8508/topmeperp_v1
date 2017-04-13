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
            List<topmeperp.Models.TND_SUPPLIER> lstSupplier = SearchSupplierByName(Request["textSupplierName"]);
            ViewBag.SearchResult = "共取得" + lstSupplier.Count + "筆資料";
            return View("Index", lstSupplier);
        }

        private List<topmeperp.Models.TND_SUPPLIER> SearchSupplierByName(string suppliername)
        {
            if (suppliername != null)
            {
                log.Info("search supplier by 名稱 =" + suppliername);
                List<topmeperp.Models.TND_SUPPLIER> lstSupplier = new List<TND_SUPPLIER>();
                using (var context = new topmepEntities())
                {
                    lstSupplier = context.TND_SUPPLIER.SqlQuery("select * from TND_SUPPLIER s "
                        + "where s.COMPANY_NAME Like '%' + @suppliername + '%';",
                         new SqlParameter("suppliername", suppliername)).ToList();
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
            TND_SUPPLIER s = null;
            if (null != id)
            {
                SupplierManage supplierService = new SupplierManage();
                s = supplierService.getSupplierById(id);
            }
            return View(s);
        }
        [HttpPost]
        public ActionResult Create(TND_SUPPLIER sup)
        {
            log.Info("create supplier process! supplier =" + sup.ToString());
            SupplierManage supplierService = new SupplierManage();
            SYS_USER u = (SYS_USER)Session["user"];
            string message = "";
            //1.更新或新增專案基本資料
            if (sup.SUPPLIER_ID == "" || sup.SUPPLIER_ID == null)
            {
                //新增供應商
                supplierService.newSupplier(sup);
                message = "新增一供應商 : 供應商編號" + sup.SUPPLIER_ID + "，其供應商名稱為" + sup.COMPANY_NAME;
            }
            else
            {
                //修改供應商基本資料
                supplierService.updateSupplier(sup);
                message = "供應商基本資料修改成功 : 其供應商編號" + sup.SUPPLIER_ID + "，供應商名稱為" + sup.COMPANY_NAME;
            }
            ViewBag.result = message;
            return View(sup);
        }
    }
}
