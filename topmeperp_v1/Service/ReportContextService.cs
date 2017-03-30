using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Migrations;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class RptForDiffProject: ContextService
    {
        public RptForDiffProject()
        {
        }
        //依據新標單資料，取得相關專案的舊報價資料
        public void RtpGetPriceFromExistProject(string newProjectId,string oldProjectId)
        {
            string sql = "SELECT DISTINCT * FROM ("
                + "SELECT newTable.EXCEL_ROW_ID ,newTable.PROJECT_ITEM_ID nid, newTable.ITEM_DESC,oldTable.ITEM_UNIT_PRICE "
                + "FROM(select * from tnd_project_item where project_id = @newProjectId and ITEM_DESC is not null) newTable "
                + " RIGHT outer JOIN "
                + "(select * from tnd_project_item where project_id = @oldProjectId  and ITEM_DESC is not null) oldTable "
                + " ON newTable.ITEM_DESC = oldTable.ITEM_DESC ) R  Where r.ITEM_DESC is not null Order BY EXCEL_ROW_ID;";
        }
    }
}
