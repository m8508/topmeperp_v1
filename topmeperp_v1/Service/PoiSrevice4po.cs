using ICSharpCode.SharpZipLib.Zip;
using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class PlanItemFromExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public IWorkbook hssfworkbook;
        public ISheet sheet = null;
        string fileformat = "xlsx";
        string projId = null;
        public List<PLAN_ITEM> lstPlanItem = null;
        public string errorMessage = null;
        //test conflicts
        public PlanItemFromExcel()
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

        public XSSFCellStyle getContentStyle()
        {
            XSSFCellStyle oStyle = (XSSFCellStyle)hssfworkbook.CreateCellStyle();

            //設定上下左右的框線
            oStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;//粗
            oStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;//細實線
            oStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//虛線
            oStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;//...  
            return oStyle;
        }
    }

    #region 預算下載表格格式處理區段
    public class BudgetFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string budgetFile = ContextService.strUploadPath + "\\budget_form.xlsx";
        string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放預算資料
        public TND_PROJECT project = null;
        public List<PLAN_BUDGET> typecodeItems = null;
        public string errorMessage = null;
        string projId = null;
        //建立預算下載表格
        public string exportExcel(TND_PROJECT project, List<PLAN_BUDGET> typecodeItems)
        {
            //1.讀取預算表格檔案
            InitializeWorkbook(budgetFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("預算");

            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(project.PROJECT_NAME);//專案名稱

            //2.建立空白欄位
            int idxRow = 4;
            IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
            //預算//預算單列小計(填入公式)
            for (int iTmp = 0; iTmp < 50; iTmp++)
            {
                row.CreateCell(8).CellFormula = "(G" + (idxRow + 1) + "*H" + (idxRow + 1) + "/100";
                row.CreateCell(iTmp);
            }
            idxRow++;

            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            fileLocation = outputPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "_預算.xlsx";
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public BudgetFormToExcel()
        {
        }
        public void InitializeWorkbook(string path)
        {
            //read the wage file via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
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
        #region 預算資料轉換 
        /**
         * 取得預算Sheet 資料
         * */
        public List<PLAN_BUDGET> ConvertDataForBudget(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":預算");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("預算");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":預算");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("預算");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有預算資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[預算]資料");
            }
            return ConverData2Budget();
        }
        /**
         * 轉換預算資料檔:預算
         * */
        protected List<PLAN_BUDGET> ConverData2Budget()
        {
            IRow row = null;
            List<PLAN_BUDGET> lstBudget = new List<PLAN_BUDGET>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            while (iRowIndex < (4))
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
                int i = 0;
                string slog = "";
                for (i = 0; i < row.Cells.Count; i++)
                {
                    slog = slog + "," + row.Cells[i];

                }
                logger.Debug("Excel Value:" + slog);
                //將各Row 資料寫入物件內
                //0.九宮格	1.次九宮格 2.折扣率 3.預算折扣率
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    lstBudget.Add(convertRow2PlanBudget(row, iRowIndex));
                }
                else
                {
                    logErrorMessage("Step1 ;取得預算資料:" + typecodeItems.Count + "筆");
                    logger.Info("Finish convert Job : count=" + typecodeItems.Count);
                    return lstBudget;
                }
                iRowIndex++;
            }
            logger.Info("Plan_Budget Count:" + iRowIndex);
            return lstBudget;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private PLAN_BUDGET convertRow2PlanBudget(IRow row, int excelrow)
        {
            PLAN_BUDGET item = new PLAN_BUDGET();
            item.PROJECT_ID = projId;
            if (row.Cells[0].ToString().Trim() != "")//0.九宮格
            {
                item.TYPE_CODE_1 = row.Cells[0].ToString();
            }
            if (row.Cells[1].ToString().Trim() != "")//1.次九宮格
            {
                item.TYPE_CODE_2 = row.Cells[1].ToString();
            }
            if (row.Cells[2].ToString().Trim() != "")//3.主系統
            {
                item.SYSTEM_MAIN = row.Cells[2].ToString();
            }
            if (row.Cells[3].ToString().Trim() != "")//4.次系統
            {
                item.SYSTEM_SUB = row.Cells[3].ToString();
            }
            //if (null != row.Cells[row.Cells.Count - 4].ToString().Trim() || row.Cells[row.Cells.Count - 4].ToString().Trim() != "")//2.投標折數
            //{
            //    try
            //    {
            //        decimal dQty = decimal.Parse(row.Cells[row.Cells.Count - 4].ToString());
            //        logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[row.Cells.Count - 4].ToString());
            //        item.TND_RATIO = dQty;
            //    }
            //    catch (Exception e)
            //    {
            //        logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[4].value=" + row.Cells[row.Cells.Count - 4].ToString());
            //        logger.Error(e);
            //    }

            //}
            if (null != row.Cells[row.Cells.Count - 2].ToString().Trim() || row.Cells[row.Cells.Count - 2].ToString().Trim() != "")//5.預算折數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[row.Cells.Count - 2].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[row.Cells.Count - 2].ToString());
                    item.BUDGET_RATIO = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[6].value=" + row.Cells[row.Cells.Count - 2].ToString());
                    logger.Error(e);
                }

            }
            item.CREATE_DATE = System.DateTime.Now;
            logger.Info("PLAN_BUDGET=" + item.ToString());
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
    #endregion

    public class PurchaseFormtoExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string templateFile = ContextService.strUploadPath + "\\Inquiry_form_template.xlsx";
        public string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放供應商報價單資料
        public PLAN_SUP_INQUIRY form = null;
        public List<PLAN_SUP_INQUIRY_ITEM> formItems = null;

        // string fileformat = "xlsx";
        //建立採購詢價單樣板
        public string exportExcel4po(PLAN_SUP_INQUIRY form, List<PLAN_SUP_INQUIRY_ITEM> formItems, bool isTemp)
        {
            //1.讀取樣板檔案
            InitializeWorkbook(templateFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            //2.填入表頭資料
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(form.PROJECT_ID);//專案名稱
            logger.Debug("Template Head_2=" + sheet.GetRow(3).Cells[0].ToString());
            sheet.GetRow(3).Cells[1].SetCellValue(form.FORM_NAME);//採購項目:
            logger.Debug("Template Head_3=" + sheet.GetRow(4).Cells[0].ToString());
            sheet.GetRow(4).Cells[1].SetCellValue(form.OWNER_NAME);//承辦人:
            logger.Debug("Template Head_4=" + sheet.GetRow(5).Cells[0].ToString());
            sheet.GetRow(5).Cells[1].SetCellValue(form.OWNER_TEL);//聯絡電話:
            logger.Debug("Template Head_5=" + sheet.GetRow(6).Cells[0].ToString());
            sheet.GetRow(6).Cells[1].SetCellValue(form.OWNER_EMAIL);//EMAIL:
            logger.Debug("Template Head_6=" + sheet.GetRow(7).Cells[0].ToString());
            sheet.GetRow(7).Cells[1].SetCellValue(form.OWNER_FAX);//FAX:
            logger.Debug("Template Head_1=" + sheet.GetRow(2).Cells[2].ToString());
            sheet.GetRow(2).Cells[3].SetCellValue(form.SUPPLIER_ID);//廠商名稱
            logger.Debug("Template Head_2=" + sheet.GetRow(3).Cells[2].ToString());
            sheet.GetRow(3).Cells[3].SetCellValue(form.CONTACT_NAME);//聯絡人
            logger.Debug("Template Head_3=" + sheet.GetRow(4).Cells[2].ToString());
            sheet.GetRow(4).Cells[3].SetCellValue(form.CONTACT_EMAIL);//電子信箱
            logger.Debug("Template Head_4=" + sheet.GetRow(5).Cells[2].ToString());
            sheet.GetRow(5).Cells[3].SetCellValue((form.DUEDATE).ToString());//報價期限
            logger.Debug("Template Head_5=" + sheet.GetRow(6).Cells[2].ToString());
            sheet.GetRow(6).Cells[3].SetCellValue(form.INQUIRY_FORM_ID);//編號

            //3.填入表單明細
            int idxRow = 9;
            foreach (PLAN_SUP_INQUIRY_ITEM item in formItems)
            {
                IRow row = sheet.GetRow(idxRow);
                //項次 項目說明    單位 數量  單價 複價  備註
                row.Cells[0].SetCellValue(idxRow - 8);///項次
                logger.Debug("Inquiry :ITEM DESC=" + item.ITEM_DESC);
                row.Cells[1].SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[2].SetCellValue(item.ITEM_UNIT);// 單位
                if (null != item.ITEM_QTY && item.ITEM_QTY.ToString().Trim() != "")
                {
                    row.Cells[3].SetCellValue(double.Parse(item.ITEM_QTY.ToString())); //數量
                }
                // row.Cells[4].SetCellValue(idxRow - 8);//單價
                // row.Cells[5].SetCellValue(idxRow - 8);複價
                row.Cells[6].SetCellValue(item.ITEM_REMARK);// 備註
                                                            //建立空白欄位
                for (int iTmp = 7; iTmp < 27; iTmp++)
                {
                    row.CreateCell(iTmp);
                }
                //填入標單項次編號 PROJECT_ITEM_ID
                row.Cells[26].SetCellValue(item.PLAN_ITEM_ID);
                idxRow++;
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用)
            string fileLocation = null;
            if (isTemp)
            {
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\Temp\\" + form.FORM_NAME + "_空白.xlsx";
            }
            else
            {
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_NAME + "_空白.xlsx";
            }
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
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
        public void convertInquiry2Plan(string fileExcel, string projectid, string iswage)
        {
            //1.讀取供應商報價單\
            InitializeWorkbook(fileExcel);
            //2.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟廠商報價單Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat);
                sheet = (HSSFSheet)hssfworkbook.GetSheet("詢價單");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            }
            //3,讀取Sheet (預設詢價單，否則抓第一張)
            if (null == sheet)
            {
                sheet = (XSSFSheet)hssfworkbook.GetSheetAt(0);
            }
            //3.讀取檔頭 資料
            //專案名稱
            form = new PLAN_SUP_INQUIRY();
            //專案名稱:	P0120
            logger.Debug(sheet.GetRow(2).Cells[0].ToString() + "," + sheet.GetRow(2).Cells[1]);
            form.PROJECT_ID = projectid;
            //工資報價單標記
            form.ISWAGE = iswage;
            //廠商名稱:	Supplier
            logger.Debug(sheet.GetRow(2).Cells[2].ToString() + "," + sheet.GetRow(2).Cells[3]);
            form.SUPPLIER_ID = sheet.GetRow(2).Cells[3].ToString(); //用供應商名稱暫代供應商編號
                                                                    //採購項目:	 詢價單名稱	
            logger.Debug(sheet.GetRow(3).Cells[0].ToString() + "," + sheet.GetRow(3).Cells[1]);
            form.FORM_NAME = sheet.GetRow(3).Cells[1].ToString();
            //聯絡人:	contact
            logger.Debug(sheet.GetRow(3).Cells[2].ToString() + "," + sheet.GetRow(3).Cells[3]);
            form.CONTACT_NAME = sheet.GetRow(3).Cells[3].ToString();
            //承辦人:
            logger.Debug(sheet.GetRow(4).Cells[0].ToString() + "," + sheet.GetRow(4).Cells[1]);
            form.OWNER_NAME = sheet.GetRow(4).Cells[1].ToString();
            //電子信箱:	contact@email.com
            logger.Debug(sheet.GetRow(4).Cells[2].ToString() + "," + sheet.GetRow(4).Cells[3]);
            form.CONTACT_EMAIL = sheet.GetRow(4).Cells[3].ToString();
            //聯絡電話:	08888888				
            logger.Debug(sheet.GetRow(5).Cells[0].ToString() + "," + sheet.GetRow(5).Cells[1]);
            form.OWNER_TEL = sheet.GetRow(5).Cells[1].ToString();
            //報價期限:	2017/1/25
            try
            {
                logger.Debug(sheet.GetRow(5).Cells[2].ToString() + "," + sheet.GetRow(5).Cells[3].ToString() + "," + sheet.GetRow(5).Cells[3].CellType);
                form.DUEDATE = DateTime.Parse(sheet.GetRow(5).Cells[3].ToString());

            }
            catch (Exception ex)
            {
                logger.Error("Datetime format error: " + ex.Message);
                throw new Exception("日期格式有錯(YYYY/MM/DD");
            }
            //電子信箱:	admin@topmep
            logger.Debug(sheet.GetRow(6).Cells[0].ToString() + "," + sheet.GetRow(6).Cells[1]);
            form.OWNER_EMAIL = sheet.GetRow(6).Cells[1].ToString();
            //編號: REF - 001
            try
            {
                logger.Debug("REF_ID=" + sheet.GetRow(6).Cells[2].ToString() + "," + sheet.GetRow(6).Cells[3]);
                form.INQUIRY_FORM_ID = sheet.GetRow(6).Cells[3].ToString().Trim();
            }
            catch (Exception ex)
            {
                form.INQUIRY_FORM_ID = "";
                logger.Error("Not Reference ID:" + ex.Message);
            }

            //FAX:
            logger.Debug(sheet.GetRow(7).Cells[0].ToString());
            form.OWNER_FAX = sheet.GetRow(7).Cells[0].ToString();

            //3.取得表單明細,逐行讀取資料
            IRow row = null;
            int iRowIndex = 9; //0 表 Row 1
            bool hasMore = true;
            //循序處理每一筆資料之欄位!!
            formItems = new List<PLAN_SUP_INQUIRY_ITEM>();
            while (hasMore)
            {
                row = sheet.GetRow(iRowIndex);
                logger.Info("excel rowid=" + iRowIndex + ",cell count=" + row.Cells.Count);
                if (row.Cells.Count < 6)
                {
                    logger.Info("Row Index=" + iRowIndex + "column count has wrong" + row.Cells.Count);
                    throw new Exception("詢價單明細欄位有問題，請調整欄位相關資料(" + row.Cells.Count + ")");
                }
                else
                {
                    try
                    {
                        logger.Debug("row id=" + iRowIndex + "Cells Count=" + row.Cells.Count + ",purchase form item vllue:" + row.Cells[0].ToString() + ","
                            + row.Cells[1] + "," + row.Cells[2] + "," + row.Cells[3] + "," + ","
                            + row.Cells[4] + "," + "," + row.Cells[5] + "," + row.Cells[6] + ",plan item id=" + row.Cells[row.Cells.Count - 1]);
                        PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
                        item.ITEM_DESC = row.Cells[1].ToString();
                        item.ITEM_UNIT = row.Cells[2].ToString();
                        //標單數量
                        decimal dQty = decimal.Parse(row.Cells[3].ToString());
                        item.ITEM_QTY = dQty;

                        //報價單單價
                        decimal dUnitPrice = decimal.Parse(row.Cells[4].ToString());
                        item.ITEM_UNIT_PRICE = dUnitPrice;

                        item.ITEM_REMARK = row.Cells[6].ToString();
                        logger.Info("Plan ITEM ID=" + row.Cells[row.Cells.Count - 1].ToString());
                        item.PLAN_ITEM_ID = row.Cells[row.Cells.Count - 1].ToString();
                        formItems.Add(item);
                    }
                    catch (Exception ex)
                    {
                        logger.Error("Row Data Error:" + iRowIndex);
                        logger.Error(ex.GetType() + ":" + ex.StackTrace);
                    }
                }
                iRowIndex++;
            }
        }
    }

}