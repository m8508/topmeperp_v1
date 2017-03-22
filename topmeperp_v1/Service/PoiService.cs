
using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class ProjectItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        string projId = null;
        public List<TND_PROJECT_ITEM> lstProjectItem = null;
        public string errorMessage = null;
        //test conflicts
        public ProjectItemFromExcel()
        {
        }
        /*讀取備標Excel 檔案!!!*/
        public void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + path); if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    fileformat = "xls";
                    hssfworkbook = new HSSFWorkbook(file);
                }
                else
                {
                    logger.Debug("process excel file for office 2007");
                    hssfworkbook = new XSSFWorkbook(file);
                }
                file.Close();
            }
        }
        //處理整理後標單..
        public void ConvertDataForTenderProject(string projectId, int startrow)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId);
                sheet = (HSSFSheet)hssfworkbook.GetSheet("整理後標單");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("整理後標單");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有整理後標單資料");
            }
            ConvertExcelToTndProjectItem(startrow);
        }
        //轉換標單內容物件
        public void ConvertExcelToTndProjectItem(int startrow)
        {
            IRow row = null;
            lstProjectItem = new List<TND_PROJECT_ITEM>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..
            while (iRowIndex < (startrow - 1))
            {
                rows.MoveNext();
                iRowIndex++;
                row = (IRow)rows.Current;
                logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            int itemId = 1;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //項次,名稱,單位,數量,單價,複價,備註,九宮格,次九宮格,主系統,次系統
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstProjectItem.Add(convertRow2TndProjectItem(itemId, row, iRowIndex));
                }
                else
                {
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return;
                }
                iRowIndex++;
                itemId++;
            }
        }
        private TND_PROJECT_ITEM convertRow2TndProjectItem(int id, IRow row, int excelrow)
        {
            TND_PROJECT_ITEM projectItem = new TND_PROJECT_ITEM();
            projectItem.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//項次
            {
                projectItem.ITEM_ID = row.Cells[0].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//名稱
            {
                projectItem.ITEM_DESC = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//單位
            {
                projectItem.ITEM_UNIT = row.Cells[2].ToString();
            }
            if (row.Cells[3].ToString().Trim() != "")//數量
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[3].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[3].ToString());
                    projectItem.ITEM_QUANTITY = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }

            }
            if (row.Cells[6].ToString().Trim() != "")//備註
            {
                projectItem.ITEM_DESC = row.Cells[6].ToString();
            }
            if (row.Cells[7].ToString().Trim() != "")//九宮格
            {
                projectItem.TYPE_CODE_1 = row.Cells[7].ToString();
            }
            if (row.Cells[8].ToString().Trim() != "")//次九宮格
            {
                projectItem.TYPE_CODE_2 = row.Cells[8].ToString();
            }
            if (row.Cells[9].ToString().Trim() != "")//主系統
            {
                projectItem.SYSTEM_MAIN = row.Cells[9].ToString();
            }
            if (row.Cells[10].ToString().Trim() != "")//主系統
            {
                projectItem.SYSTEM_MAIN = row.Cells[10].ToString();
            }
            projectItem.PROJECT_ITEM_ID = projId + "-" + id;
            projectItem.EXCEL_ROW_ID = excelrow;
            projectItem.CREATE_DATE = System.DateTime.Now;

            logger.Info("TndprojectItem=" + projectItem.ToString());
            return projectItem;
        }
        #region 消防水資料轉換 
        /**
         * 取得消防水Sheet 資料
         * */
        public List<TND_MAP_FW> ConvertDataForMapFW(string projectId)
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
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防水]資料");
            }
            return ConverData2MapFW();
        }
        /**
         * 轉換圖算數量:消防水
         * */
        protected List<TND_MAP_FW> ConverData2MapFW()
        {
            IRow row = null;
            List<TND_MAP_FW> lstMapFW = new List<TND_MAP_FW>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (1))
            {
                rows.MoveNext();
                iRowIndex++;
                row = (IRow)rows.Current;
                logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
            }
            //循序處理每一筆資料之欄位!!
            iRowIndex++;
            int itemId = 1;
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstMapFW.Add(convertRow2TndMapFW(itemId, row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量消防水:" + lstProjectItem.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return lstMapFW;
                }
                iRowIndex++;
                itemId++;
            }
            return lstMapFW;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_FW convertRow2TndMapFW(int id, IRow row, int excelrow)
        {
            TND_MAP_FW item = new TND_MAP_FW();
            item.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.項次
            {
                item.EXCEL_ITEM = row.Cells[0].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//1.圖號
            {
                item.MAP_NO = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//2.棟別
            {
                item.BUILDING_NO = row.Cells[2].ToString();
            }

            if (row.Cells[3].ToString().Trim() != "")//3.一次側位置
            {
                item.PRIMARY_SIDE = row.Cells[3].ToString();
            }
            if (row.Cells[4].ToString().Trim() != "")//4.一次側名稱
            {
                item.PRIMARY_SIDE_NAME = row.Cells[4].ToString();
            }

            if (row.Cells[5].ToString().Trim() != "")//5.二次側名稱
            {
                item.SECONDARY_SIDE = row.Cells[5].ToString();
            }
            if (row.Cells[6].ToString().Trim() != "")//6.二次側位置
            {
                item.SECONDARY_SIDE_NAME = row.Cells[6].ToString();
            }

            if (row.Cells[7].ToString().Trim() != "")//7.管材名稱
            {
                item.PIPE_NAME = row.Cells[7].ToString();
            }

            if (row.Cells[8].ToString().Trim() != "")// 8管數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.PIPE_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[9].ToString().Trim() != "")// 9.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[10].ToString().Trim() != "")// 10管長度/組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }


            if (row.Cells[11].ToString().Trim() != "")// 11管總長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;

            logger.Info("TND_MAP_FW=" + item.ToString());
            return item;
        }
        #endregion
        private void logErrorMessage(string message)
        {
            if (errorMessage == null)
            {
                errorMessage = message;
            }
            else
            {
                errorMessage = errorMessage + "<br/>" + message;
            }
        }
    }
}
