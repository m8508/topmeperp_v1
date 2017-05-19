using log4net;
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
    public class RptCompareProjectPrice : ContextService
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public RptCompareProjectPrice()
        {
        }
        //依據新標單資料，取得相關專案的舊報價資料
        public List<ProjectCompareData> RtpGetPriceFromExistProject(string srcProjectId, string tarProjectId,bool hasProject,bool hasPrice)
        {
            List<ProjectCompareData> lstCompareData = null;
            try
            {
                using (var context = new topmepEntities())
                {
                    string sql = "SELECT DISTINCT SRC.PROJECT_ID SOURCE_PROJECT_ID,"
                        + "SRC.SYSTEM_MAIN SOURCE_SYSTEM_MAIN,"
                        + "SRC.SYSTEM_SUB SOURCE_SYSTEM_SUB,"
                        + "SRC.ITEM_ID SOURCE_ITEM_ID,"
                        + "SRC.ITEM_DESC SOURCE_ITEM_DESC,"
                        + "SRC.ITEM_UNIT_PRICE SRC_UNIT_PRICE,"
                        + "TAR.ITEM_ID TARGET_ITEM_ID,"
                        + "TAR.ITEM_DESC TARGET_ITEM_DESC,"
                        + "TAR.SYSTEM_MAIN TARGET_SYSTEM_MAIN,"
                        + "TAR.SYSTEM_SUB TARGET_SYSTEM_SUB,"
                        + "TAR.PROJECT_ID TARGET_PROJECT_ID,"
                        + "SRC.EXCEL_ROW_ID EXCEL_ROW_ID FROM "
                        + "(SELECT PROJECT_ID, ITEM_ID, ITEM_DESC, ITEM_UNIT_PRICE, SYSTEM_MAIN, ISNULL(SYSTEM_SUB,'*') SYSTEM_SUB, EXCEL_ROW_ID FROM TND_PROJECT_ITEM WHERE PROJECT_ID=@srcProjectId ) SRC,"
                        + "(SELECT PROJECT_ID, ITEM_ID, ITEM_DESC, SYSTEM_MAIN, ISNULL(SYSTEM_SUB,'*') SYSTEM_SUB FROM TND_PROJECT_ITEM WHERE PROJECT_ID=@tarProjectId ) TAR "
                        + "WHERE SRC.ITEM_DESC = TAR.ITEM_DESC ";

                    if (hasProject)
                    {
                        sql = sql + "AND SRC.SYSTEM_MAIN = TAR.SYSTEM_MAIN AND SRC.SYSTEM_SUB = TAR.SYSTEM_SUB  ";
                    }
                    if (hasPrice)
                    {
                        sql = sql + "AND ITEM_UNIT_PRICE is not null ";
                    }
                    sql = sql + "ORDER BY EXCEL_ROW_ID;";

                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("srcProjectId", srcProjectId ));
                    parameters.Add(new SqlParameter("tarProjectId", tarProjectId ));
                    logger.Info("SQL=" + sql);
                    lstCompareData = context.Database.SqlQuery<ProjectCompareData>(sql, parameters.ToArray()).ToList();
                    logger.Info("Get CompareData Record Count=" + lstCompareData.Count);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.StackTrace);
            }
            return lstCompareData;
        }
    }
}
