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
                logger.Debug("Excel Value:" + row.Cells[3].ToString() + row.Cells[4] + row.Cells[5]);
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
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstMapFW.Add(convertRow2TndMapFW(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量消防水:" + lstProjectItem.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return lstMapFW;
                }
                iRowIndex++;
            }
            logger.Info("MAP_FW Count:" + iRowIndex);
            return lstMapFW;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_FW convertRow2TndMapFW(IRow row, int excelrow)
        {
            TND_MAP_FW item = new TND_MAP_FW();
            item.PROJECT_ID = projId;
           // if (row.Cells[0].ToString().Trim() != "")//0.項次
            //{
              //  item.EXCEL_ITEM = row.Cells[0].ToString();
           //}
           // if (row.Cells[1].ToString().Trim() != "")//1.圖號
           // {
             //   item.MAP_NO = row.Cells[1].ToString();
            //}
            //if (row.Cells[2].ToString().Trim() != "")//2.棟別
           // {
            //    item.BUILDING_NO = row.Cells[2].ToString();
           // }

            if (row.Cells[0].ToString().Trim() != "")//3.一次側位置
            {
                item.PRIMARY_SIDE = row.Cells[3].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//4.一次側名稱
            {
                item.PRIMARY_SIDE_NAME = row.Cells[4].ToString();
            }

            if (row.Cells[2].ToString().Trim() != "")//5.二次側名稱
            {
                item.SECONDARY_SIDE = row.Cells[2].ToString();
            }
            if (row.Cells[3].ToString().Trim() != "")//6.二次側位置
            {
                item.SECONDARY_SIDE_NAME = row.Cells[3].ToString();
            }

            if (row.Cells[4].ToString().Trim() != "")//7.管材名稱
            {
                item.PIPE_NAME = row.Cells[4].ToString();
            }

            if (row.Cells[5].ToString().Trim() != "")// 8管數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[5].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[5].ToString());
                    item.PIPE_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[5].value=" + row.Cells[5].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[6].ToString().Trim() != "")// 9.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[6].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[6].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[6].value=" + row.Cells[6].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[7].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[7].ToString().Trim() != "")// 10管長度/組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[7].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[7].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[7].value=" + row.Cells[7].ToString());
                    logger.Error(e.Message);
                }
            }


            if (row.Cells[8].ToString().Trim() != "")// 11管總長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_FW=" + item.ToString());
            return item;
        }
        #endregion
        #region 消防電資料轉換 
        // * 處理消防電，進入點
        public List<TND_MAP_FP> ConvertDataForMapFP(string projectId)
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
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[消防電]資料");
            }
            return ConverData2MapFP();
        }
        /**
         * 轉換圖算數量:消防電
         * */
        protected List<TND_MAP_FP> ConverData2MapFP()
        {
            IRow row = null;
            List<TND_MAP_FP> lstMapFP = new List<TND_MAP_FP>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < 1)
            {
                rows.MoveNext();
                iRowIndex++;
                //row = (IRow)rows.Current; (註解掉, 因為這樣讀取excel時會跳過前3行, 造成資料索引錯誤)
                // logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
             }
                //循序處理每一筆資料之欄位!!
                iRowIndex++;
                while (rows.MoveNext())
                {
                    row = (IRow)rows.Current;
                    logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                    //將各Row 資料寫入物件內
                    //1.項次	2.圖號	3.棟別	4.一次側位置	5.一次側名稱	6.二次側名稱	
                    //7.二次側位置	8.線材名稱	9.條數/組	10.線組數	11.線長度/條數	12.線總長	
                    //13.管材名稱	14.管長	15.管組數	16.管總長
                    if (row.Cells[0].ToString().ToUpper() != "END")
                    {
                        lstMapFP.Add(convertRow2TndMapFP(row, iRowIndex));
                    }
                    else
                    {
                        logErrorMessage("Step1 ;取得圖算數量消防電:" + lstMapFP.Count + "筆");
                        logger.Info("Finish convert Job : count=" + lstMapFP.Count);
                        return lstMapFP;
                    }
                    iRowIndex++;
                }
                logger.Info("MAP_FP Count:" + iRowIndex);
                return lstMapFP;
            }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_FP convertRow2TndMapFP(IRow row, int excelrow)
        {
            TND_MAP_FP item = new TND_MAP_FP();
            item.PROJECT_ID = projId;
            //0.項次	1.圖號	2.棟別	3.一次側位置 4.一次側名稱	5.二次側名稱	

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
            //6.二次側位置	7.線材名稱	8.條數/組
            if (row.Cells[6].ToString().Trim() != "")//6.二次側位置
            {
                item.SECONDARY_SIDE_NAME = row.Cells[6].ToString();
            }

            if (row.Cells[7].ToString().Trim() != "")//7.線材名稱
            {
                item.WIRE_NAME = row.Cells[7].ToString();
            }

            if (row.Cells[8].ToString().Trim() != "")// 8.條數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.WIRE_QTY_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }
            //9.線組數   10.線長度 / 條數   11.線總長      
            if (row.Cells[9].ToString().Trim() != "")// 9.線組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.WIRE_SET_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[9].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[10].ToString().Trim() != "")// 10線長度 / 條數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.WIRE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[11].ToString().Trim() != "")// 11.線總長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }
            //12.管材名稱	13.管長	14.管組數	15.管總長
            if (row.Cells[12].ToString().Trim() != "")// 12.管材名稱
            {
                item.PIPE_NAME = row.Cells[12].ToString();
            }

            if (row.Cells[13].ToString().Trim() != "")//13.管長
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[13].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[13].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[13].value=" + row.Cells[13].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[14].ToString().Trim() != "")//14.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[14].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[14].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[14].value=" + row.Cells[14].ToString());
                    logger.Error(e.Message);
                }
            }
            if (row.Cells[15].ToString().Trim() != "")//15.管組數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[15].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[15].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[15].value=" + row.Cells[15].ToString());
                    logger.Error(e.Message);
                }
            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_FW=" + item.ToString());
            return item;
        }
        #endregion
        #region 給排水資料轉換 
        /**
         * 取得給排水Sheet 資料
         * */
        public List<TND_MAP_PLU> ConvertDataForMapPLU(string projectId)
        {
            projId = projectId;
            //1.依據檔案副檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":給排水");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("給排水");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":給排水");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("給排水");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[給排水]資料");
            }
            return ConverData2MapPLU();
        }
        /**
         * 轉換圖算數量:消防水
         * */
        protected List<TND_MAP_PLU> ConverData2MapPLU()
        {
            IRow row = null;
            List<TND_MAP_PLU> lstMapPLU = new List<TND_MAP_PLU>();
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
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.管材名稱	8管數/組	9管組數	10管長度/組數	11管總長
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstMapPLU.Add(convertRow2TndMapPLU(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量給排水:" + lstProjectItem.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return lstMapPLU;
                }
                iRowIndex++;
            }
            logger.Info("MAP_PLU Count:" + iRowIndex);
            return lstMapPLU;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_PLU convertRow2TndMapPLU(IRow row, int excelrow)
        {
            TND_MAP_PLU item = new TND_MAP_PLU();
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
                    item.PIPE_COUNT_SET = dQty;
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
                    item.PIPE_SET_QTY = dQty;
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
            logger.Info("TND_MAP_PLU=" + item.ToString());
            return item;
        }
        #endregion
        #region 弱電管線資料轉換 
        /**
         * 取得弱電管線Sheet 資料
         * */
        public List<TND_MAP_LCP> ConvertDataForMapLCP(string projectId)
        {
            projId = projectId;
            //1.依據檔案副檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":弱電管線");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("弱電管線");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":弱電管線");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("弱電管線");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[弱電管線]資料");
            }
            return ConverData2MapLCP();
        }
        /**
         * 轉換圖算數量:弱電管線
         * */
        protected List<TND_MAP_LCP> ConverData2MapLCP()
        {
            IRow row = null;
            List<TND_MAP_LCP> lstMapLCP = new List<TND_MAP_LCP>();
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
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.線材名稱	8.條數/組	9.線組數	10.線長度/組數	11.線總長  12.地線名稱	13.地線條數	14.地線總長  15.管材名稱1	16.管長1	17.管組數1   18.管總長1	19.管材名稱2  20.管長2	21.管組數2  22.管總長2
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstMapLCP.Add(convertRow2TndMapLCP(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量弱電管線:" + lstProjectItem.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return lstMapLCP;
                }
                iRowIndex++;
            }
            logger.Info("MAP_LCP Count:" + iRowIndex);
            return lstMapLCP;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_LCP convertRow2TndMapLCP(IRow row, int excelrow)
        {
            TND_MAP_LCP item = new TND_MAP_LCP();
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

            if (row.Cells[7].ToString().Trim() != "")//7.線材名稱
            {
                item.WIRE_NAME = row.Cells[7].ToString();
            }

            if (row.Cells[8].ToString().Trim() != "")// 8.條數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.WIRE_QTY_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[9].ToString().Trim() != "")// 9.線組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.WIRE_SET_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[10].ToString().Trim() != "")// 10.線長度/組數   
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.WIRE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }


            if (row.Cells[11].ToString().Trim() != "")// 11.線總長  
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[12].ToString().Trim() != "") //12.地線名稱
            {
                item.GROUND_WIRE_NAME = row.Cells[12].ToString();
            }

            if (row.Cells[13].ToString().Trim() != "") // 	13.地線條數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[13].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[13].ToString());
                    item.GROUND_WIRE_QTY = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[13].value=" + row.Cells[13].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[14].ToString().Trim() != "")// 14.地線總長 		  	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[14].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[14].ToString());
                    item.GROUND_WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[14].value=" + row.Cells[14].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[15].ToString().Trim() != "") //15.管材名稱1
            {
                item.PIPE_1_NAME = row.Cells[15].ToString();
            }

            if (row.Cells[16].ToString().Trim() != "")//16.管長1
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[16].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[16].ToString());
                    item.PIPE_1_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[16].value=" + row.Cells[16].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[17].ToString().Trim() != "")//17.管組數1
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[17].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[17].ToString());
                    item.PIPE_1_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[17].value=" + row.Cells[17].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[18].ToString().Trim() != "")//18.管總長1
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[18].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[18].ToString());
                    item.PIPE_1_TOTAL_LEN = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[18].value=" + row.Cells[18].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[19].ToString().Trim() != "") //19.管材名稱2
            {
                item.PIPE_2_NAME = row.Cells[19].ToString();
            }

            if (row.Cells[20].ToString().Trim() != "")//20.管長2	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[20].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[20].ToString());
                    item.PIPE_2_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[20].value=" + row.Cells[20].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[21].ToString().Trim() != "")//21.管組數2  
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[21].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[21].ToString());
                    item.PIPE_2_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[21].value=" + row.Cells[21].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[22].ToString().Trim() != "")//22.管總長2
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[22].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[22].ToString());
                    item.PIPE_2_TOTAL_LEN = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[22].value=" + row.Cells[22].ToString());
                    logger.Error(e.Message);
                }
            }

            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_LCP=" + item.ToString());
            return item;
        }
        #endregion
        #region 電氣管線資料轉換 
        /**
         * 取得電氣管線Sheet 資料
         * */
        public List<TND_MAP_PEP> ConvertDataForMapPEP(string projectId)
        {
            projId = projectId;
            //1.依據檔案副檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":電氣管線");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("電氣管線");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":電氣管線");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("電氣管線");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有整理後標單資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[電氣管線]資料");
            }
            return ConverData2MapPEP();
        }
        /**
         * 轉換圖算數量:電氣管線
         * */
        protected List<TND_MAP_PEP> ConverData2MapPEP()
        {
            IRow row = null;
            List<TND_MAP_PEP> lstMapPEP = new List<TND_MAP_PEP>();
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
            while (rows.MoveNext())
            {
                row = (IRow)rows.Current;
                logger.Debug("Excel Value:" + row.Cells[0].ToString() + row.Cells[1] + row.Cells[2]);
                //將各Row 資料寫入物件內
                //0.項次	1.圖號	2.棟別	3.一次側位置	4.一次側名稱	5.二次側名稱	6.二次側位置	7.線材名稱	8.條數/組	9.線組數	10.線長度/條數	11.線總長  12.地線名稱	13.地線條數	14.地線總長  15.管材名稱	16.管長	17.管組數   18.管總長	
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstMapPEP.Add(convertRow2TndMapPEP(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得圖算數量電氣管線:" + lstProjectItem.Count + "筆");
                    logger.Info("Finish convert Job : count=" + lstProjectItem.Count);
                    return lstMapPEP;
                }
                iRowIndex++;
            }
            logger.Info("MAP_PEP Count:" + iRowIndex);
            return lstMapPEP;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private TND_MAP_PEP convertRow2TndMapPEP(IRow row, int excelrow)
        {
            TND_MAP_PEP item = new TND_MAP_PEP();
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

            if (row.Cells[7].ToString().Trim() != "")//7.線材名稱
            {
                item.WIRE_NAME = row.Cells[7].ToString();
            }

            if (row.Cells[8].ToString().Trim() != "")// 8.條數/組
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[8].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[8].ToString());
                    item.WIRE_QTY_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[8].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[9].ToString().Trim() != "")// 9.線組數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[9].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[9].ToString());
                    item.WIRE_SET_CNT = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[9].value=" + row.Cells[9].ToString());
                    //   logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Item_Desc= " + projectItem.ITEM_DESC + ",value=" + row.Cells[3].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[10].ToString().Trim() != "")// 10.線長度/組數   
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[10].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[10].ToString());
                    item.WIRE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[10].value=" + row.Cells[10].ToString());
                    logger.Error(e.Message);
                }
            }


            if (row.Cells[11].ToString().Trim() != "")// 11.線總長  
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[11].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[11].ToString());
                    item.WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[11].value=" + row.Cells[11].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[12].ToString().Trim() != "") //12.地線名稱
            {
                item.GROUND_WIRE_NAME = row.Cells[12].ToString();
            }

            if (row.Cells[13].ToString().Trim() != "") // 	13.地線條數	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[13].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[13].ToString());
                    item.GROUND_WIRE_QTY = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[13].value=" + row.Cells[13].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[14].ToString().Trim() != "")// 14.地線總長 		  	
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[14].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[14].ToString());
                    item.GROUND_WIRE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[14].value=" + row.Cells[14].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[15].ToString().Trim() != "") //15.管材名稱1
            {
                item.PIPE_NAME = row.Cells[15].ToString();
            }

            if (row.Cells[16].ToString().Trim() != "")//16.管長1
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[16].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[16].ToString());
                    item.PIPE_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[16].value=" + row.Cells[16].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[17].ToString().Trim() != "")//17.管組數1
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[17].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[17].ToString());
                    item.PIPE_SET = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[17].value=" + row.Cells[17].ToString());
                    logger.Error(e.Message);
                }
            }

            if (row.Cells[18].ToString().Trim() != "")//18.管總長1
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[18].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[18].ToString());
                    item.PIPE_TOTAL_LENGTH = dQty;
                }
                catch (Exception e)
                {
                    logErrorMessage("data format Error on ExcelRow=" + excelrow + ",Cells[18].value=" + row.Cells[18].ToString());
                    logger.Error(e.Message);
                }
            }

            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("TND_MAP_PEP=" + item.ToString());
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
    #region 詢價單格式處理區段
    public class InquiryFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string templateFile = "D:\\VS.NET\\topmeperp_v1\\topmeperp_v1\\UploadFile\\Inquiry_form_template.xlsx";
        string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放供應商報價單資料
        public TND_PROJECT_FORM form = null;
        public List<TND_PROJECT_FORM_ITEM> formItems = null;

        // string fileformat = "xlsx";
        //建立詢價單樣板
        public void exportExcel(TND_PROJECT_FORM form, List<TND_PROJECT_FORM_ITEM> formItems)
        {
            //1.讀取樣板檔案
            InitializeWorkbook(templateFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            //2.填入表頭資料
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue (form.PROJECT_ID);//專案名稱
            logger.Debug("Template Head_2=" + sheet.GetRow(3).Cells[0].ToString());
            sheet.GetRow(3).Cells[1].SetCellValue(form.FORM_NAME);//採購項目:
            logger.Debug("Template Head_3=" + sheet.GetRow(4).Cells[0].ToString());
            sheet.GetRow(4).Cells[1].SetCellValue (form.OWNER_NAME) ;//承辦人:
            logger.Debug("Template Head_4=" + sheet.GetRow(5).Cells[0].ToString());
            sheet.GetRow(5).Cells[1].SetCellValue (form.OWNER_TEL);//聯絡電話:
            logger.Debug("Template Head_5=" + sheet.GetRow(6).Cells[0].ToString());
            sheet.GetRow(6).Cells[1].SetCellValue(form.OWNER_EMAIL);//EMAIL:
            logger.Debug("Template Head_6=" + sheet.GetRow(7).Cells[0].ToString());
            sheet.GetRow(7).Cells[1].SetCellValue(form.OWNER_FAX);//FAX:

            //3.填入表單明細
            int idxRow = 9;
            foreach (TND_PROJECT_FORM_ITEM item in formItems)
            {
                IRow row = sheet.GetRow(idxRow);
                //項次 項目說明    單位 數量  單價 複價  備註
                row.Cells[0].SetCellValue(idxRow - 8);///項次
                logger.Debug("Inquiry :ITEM DESC="+ item.ITEM_DESC);
                row.Cells[1].SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[2].SetCellValue(item.ITEM_UNIT );// 單位
                if (null != item.ITEM_QTY && item.ITEM_QTY.ToString().Trim() != "")
                {
                    row.Cells[3].SetCellValue(double.Parse(item.ITEM_QTY.ToString())); //數量
                }
                // row.Cells[4].SetCellValue(idxRow - 8);//單價
                // row.Cells[5].SetCellValue(idxRow - 8);複價
                row.Cells[6].SetCellValue(item.ITEM_REMARK);// 備註
                //建立空白欄位
                for (int iTmp=7;iTmp < 27; iTmp++)
                {
                    row.CreateCell(iTmp);
                }
                //填入標單項次編號 PROJECT_ITEM_ID
                row.Cells[26].SetCellValue(item.PROJECT_ITEM_ID);
                idxRow++;
            }
            //4.令存新檔至專案所屬目錄
            var file = new FileStream(outputPath+"\\"+form.PROJECT_ID + "\\"+ form.FORM_ID +".xlsx", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }
        private void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                logger.Info("Read Excel File:" + path); if (file.Name.EndsWith(".xls"))
                {
                    logger.Debug("process excel file for office 2003");
                    //fileformat = "xls";
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
        public void convertInquiry2Project(string fileExcel, string projectid)
        {
            //1.讀取供應商報價單\
            InitializeWorkbook(fileExcel);
            //2,讀取Sheet (預設詢價單，否則抓第一張
            sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            if (null == sheet)
            {
                sheet = (XSSFSheet)hssfworkbook.GetSheetAt(0);
            }
            //3.讀取檔頭 資料
            //專案名稱
            form = new TND_PROJECT_FORM();
            form.PROJECT_ID = projectid;
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[0].ToString() + "," + sheet.GetRow(2).Cells[1]);
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[2].ToString() + "," + sheet.GetRow(2).Cells[3]);
            //採購項目:
            logger.Debug("Template Head_1=" + sheet.GetRow(3).Cells[0].ToString() + ",採購項目:" + sheet.GetRow(3).Cells[1]);
            logger.Debug("Template Head_1=" + sheet.GetRow(3).Cells[2].ToString() + ",聯絡人:" + sheet.GetRow(3).Cells[3]);
            //承辦人:
            // logger.Debug("Template Head_1=" + sheet.GetRow(4).Cells[0].ToString() + ",承辦人:" + sheet.GetRow(4).Cells[1]);
            //聯絡電話:
            // logger.Debug("Template Head_1=" + sheet.GetRow(5).Cells[0].ToString() + ",聯絡電話:" + sheet.GetRow(5).Cells[1]);
            //EMAIL:
            //  logger.Debug("Template Head_1=" + sheet.GetRow(6).Cells[0].ToString() + ",EMAIL:" + sheet.GetRow(6).Cells[1]);
            //FAX:
            //  logger.Debug("Template Head_1=" + sheet.GetRow(7).Cells[0].ToString() + ",FAX:" + sheet.GetRow(7).Cells[1]);

            //3.取得表單明細
            //int idxRow = 9;
            //foreach (TND_PROJECT_FORM_ITEM item in formItems)
            //{
            //    //IRow row = sheet.GetRow(idxRow);
            //    ////項次 項目說明    單位 數量  單價 複價  備註
            //    //row.Cells[0].SetCellValue(idxRow - 8);///項次
            //    //logger.Debug("Inquiry :ITEM DESC=" + item.ITEM_DESC);
            //    //row.Cells[1].SetCellValue(item.ITEM_DESC);//項目說明
            //    //row.Cells[2].SetCellValue(item.ITEM_UNIT);// 單位
            //    //if (null != item.ITEM_QTY && item.ITEM_QTY.ToString().Trim() != "")
            //    //{
            //    //    row.Cells[3].SetCellValue(double.Parse(item.ITEM_QTY.ToString())); //數量
            //    //}
            //    //// row.Cells[4].SetCellValue(idxRow - 8);//單價
            //    //// row.Cells[5].SetCellValue(idxRow - 8);複價
            //    //row.Cells[6].SetCellValue(item.ITEM_REMARK);// 備註
            //    ////建立空白欄位
            //    //for (int iTmp = 7; iTmp < 27; iTmp++)
            //    //{
            //    //    row.CreateCell(iTmp);
            //    //}
            //    ////填入標單項次編號 PROJECT_ITEM_ID
            //    //row.Cells[26].SetCellValue(item.PROJECT_ITEM_ID);
            //    idxRow++;
            //}

        }
    }
    #endregion
}
