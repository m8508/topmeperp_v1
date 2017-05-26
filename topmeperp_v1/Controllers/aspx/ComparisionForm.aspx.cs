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
        string htmlString = null;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //傳入查詢條件
                log.Info("start project id=" + Request["id"] + ",TypeCode1=" + Request["typeCode1"] + ",typecode2=" + Request["typeCode2"] + ",SystemMain=" + Request["SystemMain"] + ",Sytem Sub=" + Request["SystemSub"]);
                //取得備標品項與詢價資料
                DataTable dt = service.getComparisonDataToPivot(Request["id"], Request["typeCode1"], Request["typeCode2"], Request["SystemMain"], Request["SystemSub"],"N");
                labelMsg.Text = "共" + dt.Rows.Count + "筆";
                //grdRawData.DataSource = dt;
                //grdRawData.DataBind();
                htmlString = "<table class='table table-bordered'><tr>";
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    log.Debug("column name=" + dt.Columns[i].ColumnName);
                    htmlString = htmlString + "<th>" + dt.Columns[i].ColumnName + "</th>";
                }
                htmlString = htmlString + "</tr>";
                foreach (DataRow dr in dt.Rows)
                {
                    htmlString = htmlString + "<tr>";
                    for (int i = 0; i < dt.Columns.Count; i++) {
                        htmlString = htmlString + "<td>" + dr[i] + "</td>";
                    }
                    htmlString = htmlString + "</tr>";
                }
                htmlString = htmlString + "</table>";
                Response.Write(htmlString);
            }
        }
    }
}