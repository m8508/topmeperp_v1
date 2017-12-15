using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    /// <summary>
    /// 解析專案任務與圖算匯入
    /// </summary>
    public class ProjectTask2MapService : ProjectItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public new string fileformat = "xlsx";
        public List<PLAN_TASK2MAPITEM> lstTask2Map = null;
        public ProjectTask2MapService()
        {
        }
        public void transAllSheet(string projectId)
        {
            //消防水
            lstTask2Map = new List<PLAN_TASK2MAPITEM>();
            try
            {
                logger.Debug("Process MapFW");
                lstTask2Map.AddRange(ConvertDataForMapFW(projectId));
                errorMessage = "消防水-匯入成功";
            }
            catch(Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                errorMessage = ex.Message;
            }
            //消防電
            try
            {
                logger.Debug("Process MapFP");
                lstTask2Map.AddRange(ConvertDataForMapFP(projectId));            
                errorMessage = errorMessage + ",消防電-匯入成功";
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message + ":" + ex.StackTrace);
                errorMessage = errorMessage+ "," + ex.Message;
            }
        }
        #region 消防水資料轉換 
        /**
         * 取得消防水Sheet 資料
         * */
        public new List<PLAN_TASK2MAPITEM> ConvertDataForMapFW(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":消防水");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("消防水");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":消防水");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("消防水");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有消防水資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防水]資料");
            }
            return ConverData2MapFW();
        }
        /**
         * 轉換圖算數量:消防水
         * */
        protected new List<PLAN_TASK2MAPITEM> ConverData2MapFW()
        {
            IRow row = null;
            List<PLAN_TASK2MAPITEM> lstTask2MapFw = new List<PLAN_TASK2MAPITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current;
                //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstTask2MapFw.Add(convertRow2TndMapFW(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得專案任務之消防水品項:" + lstTask2Map.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstTask2Map.Count);
                    return lstTask2MapFw;
                }
                iRowIndex++;
            }
            logger.Info("MAP_FW Count:" + iRowIndex);
            return lstTask2Map;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件(消防水)
         * */
        private PLAN_TASK2MAPITEM convertRow2TndMapFW(IRow row, int excelrow)
        {
            PLAN_TASK2MAPITEM item = new PLAN_TASK2MAPITEM();
            item.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.專案認識識別碼
            {
                item.PRJ_UID = int.Parse(row.Cells[0].ToString());
            }
            ////if (row.Cells[1].ToString().Trim() != "")//1.棟別
            ////{
            ////    logger.Debug("cell[1]" + row.Cells[1].ToString());
            ////}
            ////if (row.Cells[2].ToString().Trim() != "")//2.一次側位置
            ////{
            ////    logger.Debug("cell[2]" + row.Cells[2].ToString());
            ////}

            ////if (row.Cells[3].ToString().Trim() != "")//3.一次側名稱
            ////{
            ////    logger.Debug("cell[3]" + row.Cells[3].ToString());
            ////}
            ////if (row.Cells[4].ToString().Trim() != "")//4.二次側名稱
            ////{
            ////    logger.Debug("cell[4]" + row.Cells[4].ToString());
            ////}

            ////if (row.Cells[5].ToString().Trim() != "")//5.二次側位置
            ////{
            ////    logger.Debug("cell[5]" + row.Cells[5].ToString());
            ////}
            if (row.Cells[8].ToString().Trim() != "")//8.二次側位置
            {
                item.PROJECT_ITEM_ID = row.Cells[8].ToString();
            }
            item.MAP_TYPE = "TND_MAP_FW";
            item.MAP_PK = 0;
            //7 以後忽略
            logger.Info("PLAN_TASK2MAPITEM : TND_MAP_FW=" + item.ToString());
            return item;
        }
        #endregion
        #region 消防電資料轉換 
        /**
         ** 取得消防電Sheet 資料
         */
        public new List<PLAN_TASK2MAPITEM> ConvertDataForMapFP(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":消防電");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("消防電");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":消防電");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("消防電");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有消防水資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防電]資料");
            }
            return ConverData2MapFP();
        }
        /**
         * 轉換圖算數量:消防水
         * */
        protected new List<PLAN_TASK2MAPITEM> ConverData2MapFP()
        {
            IRow row = null;
            List<PLAN_TASK2MAPITEM> lstTask2MapFp = new List<PLAN_TASK2MAPITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                if (row.Cells.Count != 0 && row.Cells[0].ToString().ToUpper() != "END")
                {
                    logger.Debug("FP iRows=" + iRowIndex);
                    lstTask2MapFp.AddRange(convertRow2TndMapFP(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得專案任務之消防電品項:" + lstTask2Map.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstTask2Map.Count);
                    return lstTask2MapFp;
                }
                iRowIndex++;
            }
            logger.Info("MAP_FP Count:" + iRowIndex);
            return lstTask2MapFp;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件(消防電)
         * */
        private List<PLAN_TASK2MAPITEM> convertRow2TndMapFP(IRow row, int excelrow)
        {
            //消防電含有一筆管、一筆線資料
            List<PLAN_TASK2MAPITEM> lstItem = new List<PLAN_TASK2MAPITEM>();
            PLAN_TASK2MAPITEM itemLine = new PLAN_TASK2MAPITEM();
            PLAN_TASK2MAPITEM itemPipe = new PLAN_TASK2MAPITEM();
            itemLine.PROJECT_ID = projId;
            itemPipe.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.專案認識識別碼
            {
                itemLine.PRJ_UID = int.Parse(row.Cells[0].ToString());
                itemPipe.PRJ_UID = int.Parse(row.Cells[0].ToString());
            }

            if (row.Cells[8].ToString().Trim() != "")//8
            {
                itemLine.PROJECT_ITEM_ID = row.Cells[8].ToString();
                logger.Info("PLAN_TASK2MAPITEM : TND_MAP_FP_Line=" + itemLine.PROJECT_ITEM_ID);
            }
            if (row.Cells[13].ToString().Trim() != "")//8
            {
                itemPipe.PROJECT_ITEM_ID = row.Cells[13].ToString();
                logger.Info("PLAN_TASK2MAPITEM : TND_MAP_FP_PIPE=" + itemPipe.PROJECT_ITEM_ID);
            }
            itemLine.MAP_TYPE = "TND_MAP_FP";
            itemPipe.MAP_TYPE = "TND_MAP_FP";
            itemLine.MAP_PK = 0;
            itemPipe.MAP_PK = 0;
            if (null!=itemLine.PROJECT_ITEM_ID && itemLine.PROJECT_ITEM_ID.Trim() != "")
            {
                logger.Warn("itemLine lost");
                lstItem.Add(itemLine);
            }
            if (null != itemPipe.PROJECT_ITEM_ID && itemPipe.PROJECT_ITEM_ID.Trim() != "")
            {
                logger.Warn("itemPipe lost");
                lstItem.Add(itemPipe);
            }

            return lstItem;
        }
        #endregion
    }
}