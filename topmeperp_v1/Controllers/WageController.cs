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
        // GET: Wage
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(FormCollection form)
        {
            log.Info("get project item :projectid=" + form["projectid"]);
            service.getProjectId(form["projectid"]);
            WageFormToExcel poi = new WageFormToExcel();
            poi.exportExcel(service.wageTable, service.wageTableItem);
            return View();
        }

        // GET: Wage/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: Wage/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Wage/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
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
