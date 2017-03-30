using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using topmeperp.Service;


namespace topmeperp.Controllers
{
    public class WageController : Controller
    {
        ILog log = log4net.LogManager.GetLogger(typeof(WageController));
        WageTableService service = new WageTableService();
        // GET: 下載工率excel表格
        public ActionResult Index(string id, FormCollection form)
        {
            log.Info("get project item :projectid=" + id);
            ViewBag.projectid = id;
            service.getProjectId(id);
            WageFormToExcel poi = new WageFormToExcel();
            poi.exportExcel(service.wageTable, service.wageTableItem);
            return View();
        }
        //上傳工率
        [HttpPost]
        public ActionResult uploadWageTable(string id)
        {
            log.Info("upload wage table for projectid=" + id);
            ViewBag.projectid = id;
            return View();
        }

        
        // GET: Wage/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: Wage/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Wage/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Wage/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
