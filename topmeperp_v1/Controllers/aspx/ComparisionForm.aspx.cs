using log4net;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using topmeperp.Service;

namespace topmeperp.Views.Inquiry
{
    public partial class ComparisionForm : System.Web.UI.Page
    {
        ILog log = log4net.LogManager.GetLogger(typeof(ComparisionForm));
        InquiryFormService service = new InquiryFormService();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //傳入查詢條件
                log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"]);
                //取得備標品項與詢價資料
                DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"]);
                labelMsg.Text = "共" + dt.Rows.Count + "筆";
                grdRawData.DataSource = dt;
                grdRawData.DataBind();
            }
        }

        //Binds all the GridView used in the page.//
        //private void BindGridView()
        //{

        //    //      DataTable dt = SqlLayer.GetDataTable("GetEmployee");
        //    //     Pivot pvt = new Pivot(dt);
        //    DataTable dt = service.getComparisonDataToPivot("P0120", null, null, null, null);
        //    labelMsg.Text = "共" + dt.Rows.Count + "筆";
        //    grdRawData.DataSource = dt;
        //    grdRawData.DataBind();

        //    //   grdCompanyYear.DataSource = pvt.PivotData("Company", "CTC", AggregateFunction.Count, "Year");
        //    //   grdCompanyYear.DataBind();

        //}
    }
}