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
        CostAnalysisDataService service = new CostAnalysisDataService();
        public TND_PROJECT project = null;
        public List<DirectCost> typecodeItems = null;
        public string errorMessage = null;
        string projId = null;

        //建立預算下載表格
        public string exportExcel(TND_PROJECT project)
        {
            List<DirectCost> typecodeItems = service.getDirectCost4Budget(project.PROJECT_ID);
            //1.讀取預算表格檔案
            InitializeWorkbook(budgetFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("預算");

            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(project.PROJECT_NAME);//專案名稱
            //3.填入資料
            int idxRow = 4;
            foreach (DirectCost item in typecodeItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //主九宮格編碼、次九宮格編碼、分項名稱(成本價)、合約金額、材料成本、預算折扣率、預算金額
                //主九宮格編碼
                row.CreateCell(0).SetCellValue(item.MAINCODE);
                //次九宮格編碼
                if (null != item.SUB_CODE && item.SUB_CODE.ToString().Trim() != "")
                {
                    row.CreateCell(1).SetCellValue(double.Parse(item.SUB_CODE.ToString()));
                }
                else
                {
                    row.CreateCell(1).SetCellValue("");
                }
                //分項名稱
                logger.Debug("ITEM DESC=" + item.MAINCODE_DESC);
                row.CreateCell(2).SetCellValue(item.MAINCODE_DESC + "-" + item.SUB_DESC);
                //合約金額
                if (null != item.CONTRACT_PRICE && item.CONTRACT_PRICE.ToString().Trim() != "")
                {
                    row.CreateCell(3).SetCellValue(double.Parse(item.CONTRACT_PRICE.ToString()));
                }
                else
                {
                    row.CreateCell(3).SetCellValue("");
                }
                //材料成本
                if (null != item.MATERIAL_COST_INMAP && item.MATERIAL_COST_INMAP.ToString().Trim() != "")
                {
                    row.CreateCell(4).SetCellValue(double.Parse(item.MATERIAL_COST_INMAP.ToString()));
                }
                else
                {
                    row.CreateCell(4).SetCellValue("");
                }
                //材料折扣率 
                row.CreateCell(5).SetCellValue("");
                //材料預算
                ICell cel6 = row.CreateCell(6);
                cel6.CellFormula = "(E" + (idxRow + 1) + "*F" + (idxRow + 1) + "/100)";
                cel6.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                //工資成本
                row.CreateCell(7).SetCellValue("");  //圖算*工率
                if (null != item.MAN_DAY_4EXCEL && item.MAN_DAY_4EXCEL.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellFormula(item.MAN_DAY_4EXCEL.ToString() + "*H3");
                }
                //工資折扣率 
                row.CreateCell(8).SetCellValue("");
                foreach (ICell c in row.Cells)
                {
                    c.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                }
                //工資預算
                ICell cel9 = row.CreateCell(9);
                cel9.CellFormula = "(H" + (idxRow + 1) + "*I" + (idxRow + 1) + "/100)";
                cel9.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                //預算金額
                ICell cel10 = row.CreateCell(10);
                cel10.CellFormula = "(E" + (idxRow + 1) + "*F" + (idxRow + 1) + "/100)+(H" + (idxRow + 1) + "*I" + (idxRow + 1) + "/100)";
                cel10.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                logger.Debug("getBudget cell style rowid=" + idxRow);
                idxRow++;
            }
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
                //0.九宮格	1.次九宮格 2.主系統 3.次系統 4.預算折扣率 6.工資預算折扣率
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
            //if (row.Cells[2].ToString().Trim() != "")//3.主系統
            //{
            // item.SYSTEM_MAIN = row.Cells[2].ToString();
            //}
            //if (row.Cells[3].ToString().Trim() != "")//4.次系統
            //{
            //item.SYSTEM_SUB = row.Cells[3].ToString();
            //}
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
            if (null != row.Cells[row.Cells.Count - 6].ToString().Trim() || row.Cells[row.Cells.Count - 6].ToString().Trim() != "")//5.預算折數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[row.Cells.Count - 6].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[row.Cells.Count - 6].ToString());
                    item.BUDGET_RATIO = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[6].value=" + row.Cells[row.Cells.Count - 6].ToString());
                    logger.Error(e);
                }

            }
            if (null != row.Cells[row.Cells.Count - 3].ToString().Trim() || row.Cells[row.Cells.Count - 3].ToString().Trim() != "")//6.工資預算折數
            {
                try
                {
                    decimal dQty = decimal.Parse(row.Cells[row.Cells.Count - 3].ToString());
                    logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[row.Cells.Count - 3].ToString());
                    item.BUDGET_WAGE_RATIO = dQty;
                }
                catch (Exception e)
                {
                    logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[8].value=" + row.Cells[row.Cells.Count - 3].ToString());
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
        string templateFile4All = ContextService.strUploadPath + "\\Inquiry_form_templateAll.xlsx";
        public string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        XSSFCellStyle style = null;
        XSSFCellStyle styleNumber = null;
        //存放供應商報價單資料
        public PLAN_SUP_INQUIRY form = null;
        public List<PLAN_SUP_INQUIRY_ITEM> formItems = null;

        // string fileformat = "xlsx";
        //建立採購詢價單樣板
        public string exportExcel4po(PLAN_SUP_INQUIRY form, List<PLAN_SUP_INQUIRY_ITEM> formItems, bool isTemp, bool isReal)
        {
            //1.讀取樣板檔案
            InitializeWorkbook(templateFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            //2.填入表頭資料
            InitialInquiryForm(form);
            //3.填入表單明細
            int idxRow = 9;
            foreach (PLAN_SUP_INQUIRY_ITEM item in formItems)
            {
                IRow row = sheet.GetRow(idxRow);
                //項次 項目說明    單位 數量  單價 複價  備註
                row.Cells[0].SetCellValue(item.ITEM_ID);///項次
                logger.Debug("Inquiry :ITEM DESC=" + item.ITEM_DESC);
                row.Cells[1].SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[2].SetCellValue(item.ITEM_UNIT);// 單位
                if (null != item.ITEM_QTY && item.ITEM_QTY.ToString().Trim() != "")
                {
                    row.Cells[3].SetCellValue(double.Parse(item.ITEM_QTY.ToString())); //數量
                }
                if (isReal && null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString() != "")
                {
                    row.Cells[4].SetCellValue(item.ITEM_UNIT_PRICE.ToString());
                    row.Cells[5].SetCellFormula("D" + (idxRow + 1) + "*E" + (idxRow + 1));//複價
                }

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
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\Temp\\" + form.FORM_NAME + "\\" + form.FORM_NAME + "_空白.xlsx";
            }
            else
            {
                if (isReal)
                {
                    fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_NAME + ".xlsx";
                }
                else
                {
                    fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_NAME + "_空白.xlsx";
                }
            }
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        // string fileformat = "xlsx";
        //建立採購詢價單樣板
        public string exportExcel4poAll(PLAN_SUP_INQUIRY form, List<PLAN_SUP_INQUIRY_ITEM> formItems, bool isTemp, bool isReal)
        {
            //1.讀取樣板檔案
            InitializeWorkbook(templateFile4All);
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("詢價單");
            InitialInquiryForm(form);

            //3.填入表單明細
            int idxRow = 9;
            foreach (PLAN_SUP_INQUIRY_ITEM item in formItems)
            {
                IRow row = sheet.CreateRow(idxRow);
                //項次 項目說明    單位 數量  材料單價 材料複價  工資單價 工資複價 備註
                row.CreateCell(0);
                row.Cells[0].SetCellValue(item.ITEM_ID);///項次
                row.Cells[0].CellStyle = style;
                logger.Debug("Inquiry :ITEM DESC=" + item.ITEM_DESC);
                row.CreateCell(1);
                row.Cells[1].SetCellValue(item.ITEM_DESC);//項目說明
                row.Cells[1].CellStyle = style;
                row.CreateCell(2);
                row.Cells[2].SetCellValue(item.ITEM_UNIT);// 單位
                row.Cells[2].CellStyle = style;
                row.CreateCell(3);
                if (null != item.ITEM_QTY && item.ITEM_QTY.ToString().Trim() != "")
                {
                    row.Cells[3].SetCellValue(double.Parse(item.ITEM_QTY.ToString())); //數量
                }
                row.Cells[3].CellStyle = style;
                row.CreateCell(4);//單價
                row.CreateCell(5);//複價
                if (isReal && item.ITEM_UNIT_PRICE != null)
                {
                    row.Cells[4].SetCellValue(item.ITEM_UNIT_PRICE.ToString());
                    row.Cells[5].SetCellFormula("D" + (idxRow + 1) + "*E" + (idxRow + 1));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//工資
                row.CreateCell(7);
                if (isReal && item.WAGE_PRICE != null)
                {
                    row.Cells[6].SetCellValue(item.ITEM_UNIT_PRICE.ToString());
                    row.Cells[7].SetCellFormula("D" + (idxRow + 1) + "*G" + (idxRow + 1));
                }
                else
                {
                    row.Cells[6].SetCellValue("");
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.Cells[7].CellStyle = styleNumber;

                row.CreateCell(8);
                row.Cells[8].SetCellValue(item.ITEM_REMARK);// 備註
                row.Cells[8].CellStyle = style;
                //建立空白
                for (int iTmp = 9; iTmp < 26; iTmp++)
                {
                    row.CreateCell(iTmp);
                }
                //填入標單項次編號 PROJECT_ITEM_ID
                row.Cells[25].SetCellValue(item.PLAN_ITEM_ID);
                idxRow++;
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用)
            string fileLocation = null;
            if (isTemp)
            {
                fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\Temp\\" + form.FORM_NAME + "[工料]_空白.xlsx";
            }
            else
            {
                if (isReal)
                {
                    fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_NAME + "[工料].xlsx";
                }
                else
                {
                    fileLocation = outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.FORM_NAME + "[工料]_空白.xlsx";
                }
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

            //2.讀取檔頭 資料
            processForm(projectid, iswage);

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
                        if (row.Cells[0].ToString().ToUpper() == "END")
                        {
                            //設定結束標記
                            hasMore = false;
                        }
                        else
                        {
                            PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
                            item.ITEM_DESC = row.Cells[1].ToString();
                            item.ITEM_UNIT = row.Cells[2].ToString();
                            //標單數量
                            try
                            {
                                decimal dQty = decimal.Parse(row.Cells[3].ToString());
                                item.ITEM_QTY = dQty;
                            }
                            catch (Exception ex)
                            {
                                logger.Warn("RowId=" + iRowIndex + " not have Qty!!" + ex.StackTrace);
                            }

                            //報價單單價
                            try
                            {
                                decimal dUnitPrice = decimal.Parse(row.Cells[4].ToString());
                                logger.Warn("RowId=" + iRowIndex +"," + item.ITEM_DESC + ", Unprice=" + dUnitPrice);
                                item.ITEM_UNIT_PRICE = dUnitPrice;
                            }
                            catch (Exception ex)
                            {
                                logger.Warn("RowId=" + iRowIndex + " not have Unit price!!" + ex.StackTrace);
                            }

                            item.ITEM_REMARK = row.Cells[6].ToString();
                            logger.Info("Plan ITEM ID=" + row.Cells[row.Cells.Count - 1].ToString());
                            item.PLAN_ITEM_ID = row.Cells[row.Cells.Count - 1].ToString();
                            formItems.Add(item);
                        }
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
        public void convertInquiryAll(string fileExcel, string projectid)
        {
            //1.讀取供應商報價單\
            InitializeWorkbook(fileExcel);
            //2.讀取檔頭 資料
            processForm(projectid, "A");

            //3.取得表單明細,逐行讀取資料
            IRow row = null;
            int iRowIndex = 9; //0 表 Row 1
            bool hasMore = true;
            //循序處理每一筆資料之欄位!!
            //項次 項目說明    單位 數量  材料單價 材料複價  工資單價 工資複價 備註
            formItems = new List<PLAN_SUP_INQUIRY_ITEM>();
            while (hasMore)
            {
                row = sheet.GetRow(iRowIndex);
                if (null == row || row.Cells.Count < 6)
                {
                    hasMore = false;
                }
                else
                {
                    logger.Info("excel rowid=" + iRowIndex + ",cell count=" + row.Cells.Count);
                    try
                    {
                        logger.Debug("row id=" + iRowIndex + "Cells Count=" + row.Cells.Count + ",purchase form item vllue:" + row.Cells[0].ToString() + ","
                            + row.Cells[1] + "," + row.Cells[2] + "," + row.Cells[3] + "," + ","
                            + row.Cells[4] + "," + "," + row.Cells[5] + "," + row.Cells[6] + ",plan item id=" + row.Cells[row.Cells.Count - 1]);
                        PLAN_SUP_INQUIRY_ITEM item = new PLAN_SUP_INQUIRY_ITEM();
                        item.ITEM_DESC = row.Cells[1].ToString();
                        item.ITEM_UNIT = row.Cells[2].ToString();
                        //標單數量
                        try
                        {
                            if (row.Cells[3].ToString() != "")
                            {
                                decimal dQty = decimal.Parse(row.Cells[3].ToString());
                                item.ITEM_QTY = dQty;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Warn("Row Index=" + iRowIndex + "item.ITEM_QTY Error:" + ex.StackTrace);
                        }
                        //材料單價
                        try
                        {
                            if (row.Cells[4].ToString() != "")
                            {
                                decimal dUnitPrice = decimal.Parse(row.Cells[4].ToString());
                                item.ITEM_UNIT_PRICE = dUnitPrice;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Warn("Row Index=" + iRowIndex + "item.ITEM_UNIT_PRICE Error:" + ex.StackTrace);
                        }
                        //材料單價
                        try
                        {
                            if (row.Cells[6].ToString() != "")
                            {
                                decimal dWagePrice = decimal.Parse(row.Cells[6].ToString());
                                item.WAGE_PRICE = dWagePrice;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Warn("Row Index=" + iRowIndex + "item.WAGE_PRICE Error:" + ex.StackTrace);
                        }

                        item.ITEM_REMARK = row.Cells[8].ToString();
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
        //處理詢價單，報價單表頭
        private void processForm(string projectid, string iswage)
        {
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
            //4.讀取檔頭 資料
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
                if (null == sheet.GetRow(5).Cells[3] || "" == sheet.GetRow(5).Cells[3].ToString())
                {
                    form.DUEDATE = DateTime.Now;
                }
                else
                {
                    string[] aryDate = sheet.GetRow(5).Cells[3].ToString().Split('/');
                    int intYear = int.Parse(aryDate[0]); ;
                    int intMonth = int.Parse(aryDate[1]);
                    int intDay = int.Parse(aryDate[2]);
                    if (intYear < 1900)
                    {
                        intYear = intYear + 1911;
                    }
                    DateTime dtDueDate = new DateTime(intYear, intMonth, intDay);
                    form.DUEDATE = dtDueDate;
                }
                logger.Debug("form.DUEDATE:" + form.DUEDATE);
            }
            catch (Exception ex)
            {
                logger.Error("Datetime format error: " + ex.Message);
                form.DUEDATE = DateTime.Now;
                // throw new Exception("日期格式有錯(YYYY/MM/DD");
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
            logger.Debug(sheet.GetRow(7).Cells[1].ToString());
            form.OWNER_FAX = sheet.GetRow(7).Cells[1].ToString();
        }
        //2.填入表頭資料
        private void InitialInquiryForm(PLAN_SUP_INQUIRY form)
        {
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
            if (form.SUPPLIER_ID != null && "" != form.SUPPLIER_ID)
            {
                sheet.GetRow(6).Cells[3].SetCellValue(form.INQUIRY_FORM_ID);//廠商詢價單編號
            }
            else
            {
                sheet.GetRow(6).Cells[3].SetCellValue("");//空白詢價單不提供編號
            }
        }
    }
    #region 公司費用預算下載表格格式處理區段
    public class ExpBudgetFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string budgetFile = ContextService.strUploadPath + "\\expense_budget_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放公司費用預算資料
        PurchaseFormService service = new PurchaseFormService();
        public List<FIN_SUBJECT> subjects = null;
        public string errorMessage = null;
        int budgetYear = 0;

        //建立公司費用預算下載表格
        public string exportExcel()
        {
            List<FIN_SUBJECT> subjects = service.getExpBudgetSubject();
            string budgetYear = null;
            if (DateTime.Now.Month > 3)
            {
                budgetYear = DateTime.Now.Year.ToString();
            }
            else
            {
                budgetYear = (DateTime.Now.Year - 1).ToString();
            }
            logger.Debug("budgetYear = " + budgetYear);
            //1.讀取公司費用預算表格檔案
            InitializeWorkbook(budgetFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("公司費用預算");

            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(budgetYear);//公司費用預算年度
            sheet.GetRow(1).Cells[13].SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));//製表日期
            //3.填入資料
            int idxRow = 4;
            foreach (FIN_SUBJECT item in subjects)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //項目、項目代碼
                //項目
                row.CreateCell(0).SetCellValue(item.SUBJECT_NAME);
                //項目代碼
                row.CreateCell(1).SetCellValue(item.FIN_SUBJECT_ID);
                row.CreateCell(2).SetCellValue("");
                row.CreateCell(3).SetCellValue("");
                row.CreateCell(4).SetCellValue("");
                row.CreateCell(5).SetCellValue("");
                row.CreateCell(6).SetCellValue("");
                row.CreateCell(7).SetCellValue("");
                row.CreateCell(8).SetCellValue("");
                row.CreateCell(9).SetCellValue("");
                row.CreateCell(10).SetCellValue("");
                row.CreateCell(11).SetCellValue("");
                row.CreateCell(12).SetCellValue("");
                row.CreateCell(13).SetCellValue("");
                foreach (ICell c in row.Cells)
                {
                    c.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                }
                ICell cel14 = row.CreateCell(14);
                cel14.CellFormula = "C" + (idxRow + 1) + "+D" + (idxRow + 1) + "+E" + (idxRow + 1) + "+F" + (idxRow + 1) + "+G" + (idxRow + 1) + "+H" + (idxRow + 1)
                + "+I" + (idxRow + 1) + "+J" + (idxRow + 1) + "+K" + (idxRow + 1) + "+L" + (idxRow + 1) + "+M" + (idxRow + 1) + "+N" + (idxRow + 1);
                cel14.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                logger.Debug("getSubject cell style rowid=" + idxRow);
                idxRow++;
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            fileLocation = outputPath + "\\" + budgetYear + "_公司費用預算.xlsx";
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public ExpBudgetFormToExcel()
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
        #region 公司費用預算資料轉換 
        /**
         * 取得公司費用預算Sheet 資料
         * */
        public List<FIN_EXPENSE_BUDGET> ConvertDataForExpBudget(int year)
        {
            budgetYear = year;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for Year=" + year + ":公司費用預算");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("公司費用預算");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for Year=" + year + ":預算");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("公司費用預算");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有公司費用預算資料(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[公司費用預算]資料");
            }
            return ConverData2ExpBudget();
        }
        /**
         * 轉換公司費用預算資料檔:公司費用預算
         * */
        protected List<FIN_EXPENSE_BUDGET> ConverData2ExpBudget()
        {
            IRow row = null;
            List<FIN_EXPENSE_BUDGET> lstExpBudget = new List<FIN_EXPENSE_BUDGET>();
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
                //0.項目代碼 2.1月金額 3.2月金額 4.3月金額 5.4月金額 6.5月金額 7.6月金額 8.7月金額 9.8月金額 10.9月金額 11.10月金額 12.11月金額 13.12月金額
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    List<FIN_EXPENSE_BUDGET> lst = convertRow2ExpBudget(row, iRowIndex);
                    foreach (FIN_EXPENSE_BUDGET it in lst)
                    {
                        lstExpBudget.Add(it);
                    }
                }
                else
                {
                    logErrorMessage("Step1 ;取得公司費用預算資料:" + subjects.Count + "筆");
                    logger.Info("Finish convert Job : count=" + subjects.Count);
                    return lstExpBudget;
                }
                iRowIndex++;
            }
            logger.Info("Expense_Budget Count:" + iRowIndex);
            return lstExpBudget;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private List<FIN_EXPENSE_BUDGET> convertRow2ExpBudget(IRow row, int excelrow)
        {
            List<FIN_EXPENSE_BUDGET> lst = new List<FIN_EXPENSE_BUDGET>();
            String strItemId = row.Cells[1].ToString();
            if (null != strItemId && strItemId != "")
            {
                for (int i = 0; i < 6; i++)
                {
                    FIN_EXPENSE_BUDGET item = new FIN_EXPENSE_BUDGET();
                    item.BUDGET_YEAR = budgetYear;
                    item.CURRENT_YEAR = budgetYear;
                    item.BUDGET_MONTH = i + 7;
                    item.SUBJECT_ID = row.Cells[1].ToString();
                    item.CREATE_DATE = System.DateTime.Now;
                    if (null != row.Cells[i + 2].ToString().Trim() || row.Cells[i + 2].ToString().Trim() != "")
                    {
                        try
                        {
                            decimal dAmt = decimal.Parse(row.Cells[i + 2].ToString());
                            logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[i + 2].ToString());
                            item.AMOUNT = dAmt;
                        }
                        catch (Exception e)
                        {
                            logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[i+2].value=" + row.Cells[i + 2].ToString());
                            logger.Error(e);
                        }
                    }
                    lst.Add(item);
                }

                for (int i = 6; i < 12; i++)
                {
                    FIN_EXPENSE_BUDGET item = new FIN_EXPENSE_BUDGET();
                    item.BUDGET_YEAR = budgetYear;
                    item.CURRENT_YEAR = budgetYear + 1;
                    item.BUDGET_MONTH = i - 5;
                    item.SUBJECT_ID = row.Cells[1].ToString();
                    item.CREATE_DATE = System.DateTime.Now;
                    if (null != row.Cells[i + 2].ToString().Trim() || row.Cells[i + 2].ToString().Trim() != "")
                    {
                        try
                        {
                            decimal dAmt = decimal.Parse(row.Cells[i + 2].ToString());
                            logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[i + 2].ToString());
                            item.AMOUNT = dAmt;
                        }
                        catch (Exception e)
                        {
                            logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[i+2].value=" + row.Cells[i + 2].ToString());
                            logger.Error(e);
                        }
                    }
                    lst.Add(item);
                }
            }
            return lst;
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
    #region 費用下載表格格式處理區段
    public class ExpenseFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string expenseFile = ContextService.strUploadPath + "\\expense_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";

        //存放費用表資料
        PurchaseFormService service = new PurchaseFormService();
        public ExpenseFormFunction ExpTable = null;
        public List<ExpenseBudgetSummary> EXPTableItem = null;
        public List<ExpenseBudgetSummary> SiteTableItem = null;
        ExpenseBudgetSummary Budget = null;
        ExpenseBudgetSummary siteBudget = null;
        ExpenseBudgetSummary ExpAmt = null;
        ExpenseBudgetSummary EarlyExpAmt = null;
        ExpenseBudgetSummary siteEarlyExpAmt = null;
        public string errorMessage = null;

        //建立預算下載表格
        public string exportExcel(ExpenseFormFunction ExpTable, List<ExpenseBudgetSummary> EXPTableItem, List<ExpenseBudgetSummary> SiteTableItem, ExpenseBudgetSummary ExpAmt, ExpenseBudgetSummary EarlyExpAmt, ExpenseBudgetSummary siteEarlyExpAmt)
        {
            int budgetYear = 0;
            if (int.Parse(ExpTable.OCCURRED_MONTH.ToString()) > 6)
            {
                budgetYear = int.Parse(ExpTable.OCCURRED_YEAR.ToString());
            }
            else
            {
                budgetYear = int.Parse((ExpTable.OCCURRED_YEAR - 1).ToString());
            }
            Budget = service.getTotalExpBudgetAmount(budgetYear);
            if (null != ExpTable.PROJECT_ID && ExpTable.PROJECT_ID != "")
            {
                siteBudget = service.getTotalSiteBudgetAmount(ExpTable.PROJECT_ID);
            }
            //1.讀取費用表格檔案
            InitializeWorkbook(expenseFile);
            if (null != ExpTable.PROJECT_ID && ExpTable.PROJECT_ID != "")
            {
                float page = 1;
                logger.Info("工地費用單頁數 =" + Convert.ToInt16(Math.Ceiling(float.Parse(EXPTableItem.Count.ToString()) / 47)));
                if (float.Parse(EXPTableItem.Count.ToString()) / 47 > 1)
                {
                    page = float.Parse(EXPTableItem.Count.ToString()) / 47;
                    sheet = (XSSFSheet)hssfworkbook.GetSheet("費用表_2頁");
                }
                else
                {
                    sheet = (XSSFSheet)hssfworkbook.GetSheet("費用表");
                }
                //2.填入表頭資料
                logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
                sheet.GetRow(0).Cells[4].SetCellValue("協成水電工程事業有限公司 工地費用單");
                sheet.GetRow(1).Cells[3].SetCellValue(ExpTable.PROJECT_NAME);//專案名稱
                sheet.GetRow(0).Cells[20].SetCellValue("第 1 頁");
                sheet.GetRow(0).Cells[22].SetCellValue("共" + Convert.ToInt16(Math.Ceiling(page)) + "頁");
                logger.Debug("Table Head_2=" + sheet.GetRow(1).Cells[9].ToString());
                sheet.GetRow(1).Cells[9].SetCellValue(ExpTable.EXP_FORM_ID);//費用單編號
                logger.Debug("Table Head_3=" + sheet.GetRow(1).Cells[14].ToString());
                sheet.GetRow(1).Cells[14].SetCellValue(ExpTable.PAYEE);//請款人
                logger.Debug("Table Head_4=" + sheet.GetRow(1).Cells[19].ToString());
                sheet.GetRow(1).Cells[19].SetCellValue(ExpTable.OCCURRED_YEAR.ToString() + "/" + ExpTable.OCCURRED_MONTH.ToString());//費用發生年月
                logger.Debug("Table Head_5=" + sheet.GetRow(1).Cells[22].ToString());
                sheet.GetRow(1).Cells[22].SetCellValue(Convert.ToDateTime(ExpTable.CREATE_DATE).ToShortDateString());//請款日期
                if (float.Parse(EXPTableItem.Count.ToString()) / 47 > 1)
                {
                    sheet.GetRow(48).Cells[4].SetCellValue("協成水電工程事業有限公司 工地費用單");
                    sheet.GetRow(48).Cells[20].SetCellValue("第 2 頁");
                    sheet.GetRow(48).Cells[22].SetCellValue("共" + Convert.ToInt16(Math.Ceiling(page)) + "頁");
                    sheet.GetRow(49).Cells[3].SetCellValue(ExpTable.PROJECT_NAME);//專案名稱
                    sheet.GetRow(49).Cells[9].SetCellValue(ExpTable.EXP_FORM_ID);//費用單編號
                    sheet.GetRow(49).Cells[14].SetCellValue(ExpTable.PAYEE);//請款人
                    sheet.GetRow(49).Cells[19].SetCellValue(ExpTable.OCCURRED_YEAR.ToString() + "/" + ExpTable.OCCURRED_MONTH.ToString());//費用發生年月
                    sheet.GetRow(49).Cells[22].SetCellValue(Convert.ToDateTime(ExpTable.CREATE_DATE).ToShortDateString());//請款日期
                }
                if (float.Parse(EXPTableItem.Count.ToString()) / 47 > 1)
                {
                    logger.Debug("Table Head_6=" + sheet.GetRow(86).Cells[1].ToString());
                    sheet.GetRow(86).Cells[1].SetCellValue(ExpTable.REMARK);//說明事項
                    logger.Debug("Table Head_7=" + sheet.GetRow(86).Cells[12].ToString());
                    if (siteBudget.TOTAL_BUDGET.ToString() != "")
                    {
                        sheet.GetRow(86).Cells[13].SetCellValue(double.Parse(siteBudget.TOTAL_BUDGET.ToString()));//總預算
                    }
                    logger.Debug("Table Head_8=" + sheet.GetRow(89).Cells[12].ToString());
                    if (siteEarlyExpAmt.AMOUNT.ToString() != "")
                    {
                        sheet.GetRow(89).Cells[13].SetCellValue(double.Parse(siteEarlyExpAmt.AMOUNT.ToString()));//前期累計
                    }
                    else
                    {
                        sheet.GetRow(43).Cells[13].SetCellValue("0");
                    }
                    logger.Debug("Table Head_9=" + sheet.GetRow(90).Cells[12].ToString());
                    sheet.GetRow(90).Cells[13].SetCellValue(double.Parse(ExpAmt.AMOUNT.ToString()));//本期金額
                    logger.Debug("Table Head_10=" + sheet.GetRow(91).Cells[12].ToString());                                                                              //工資預算
                    sheet.GetRow(91).Cells[13].CellFormula = "(N90+N91)";
                    logger.Debug("Table Head_11=" + sheet.GetRow(92).Cells[12].ToString());
                    sheet.GetRow(92).Cells[13].SetCellValue(Convert.ToDateTime(ExpTable.PAYMENT_DATE).ToShortDateString());//支付日期
                }
                else
                {
                    logger.Debug("Table Head_6=" + sheet.GetRow(40).Cells[1].ToString());
                    sheet.GetRow(40).Cells[1].SetCellValue(ExpTable.REMARK);//說明事項
                    logger.Debug("Table Head_7=" + sheet.GetRow(40).Cells[12].ToString());
                    if (siteBudget.TOTAL_BUDGET.ToString() != "")
                    {
                        sheet.GetRow(40).Cells[13].SetCellValue(double.Parse(siteBudget.TOTAL_BUDGET.ToString()));//總預算
                    }
                    logger.Debug("Table Head_8=" + sheet.GetRow(43).Cells[12].ToString());
                    if (siteEarlyExpAmt.AMOUNT.ToString() != "")
                    {
                        sheet.GetRow(43).Cells[13].SetCellValue(double.Parse(siteEarlyExpAmt.AMOUNT.ToString()));//前期累計
                    }
                    else
                    {
                        sheet.GetRow(43).Cells[13].SetCellValue("0");
                    }
                    logger.Debug("Table Head_9=" + sheet.GetRow(44).Cells[12].ToString());
                    sheet.GetRow(44).Cells[13].SetCellValue(double.Parse(ExpAmt.AMOUNT.ToString()));//本期金額
                    logger.Debug("Table Head_10=" + sheet.GetRow(45).Cells[12].ToString());                                                                              //工資預算
                    sheet.GetRow(45).Cells[13].CellFormula = "(N44+N45)";
                    logger.Debug("Table Head_11=" + sheet.GetRow(46).Cells[12].ToString());
                    sheet.GetRow(46).Cells[13].SetCellValue(Convert.ToDateTime(ExpTable.PAYMENT_DATE).ToShortDateString());//支付日期
                }
                //3.填入資料
                int idxRow = 4;
                foreach (ExpenseBudgetSummary item in SiteTableItem)
                {
                    IRow row = sheet.GetRow(idxRow);//.CreateRow(idxRow);
                    logger.Info("Row Id=" + idxRow);
                    //項次、品名/摘要、單位、預算(合約)金額、前期累計金額、本期金額、累計金額、累計金額比率、會計名稱
                    row.Cells[0].SetCellValue(item.NO);
                    //品名或摘要
                    logger.Debug("ITEM_REMARK=" + item.ITEM_REMARK);
                    row.Cells[1].SetCellValue(item.ITEM_REMARK);
                    //單位
                    row.Cells[7].SetCellValue(item.ITEM_UNIT);
                    //預算(合約)金額
                    if (null != item.BUDGET_AMOUNT && item.BUDGET_AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[9].SetCellValue(double.Parse(item.BUDGET_AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[9].SetCellValue("");
                    }
                    //前期累計金額
                    if (null != item.CUM_AMOUNT && item.CUM_AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[14].SetCellValue(double.Parse(item.CUM_AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[14].SetCellValue("");
                    }
                    //本期金額
                    if (null != item.AMOUNT && item.AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[16].SetCellValue(double.Parse(item.AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[16].SetCellValue("");
                    }
                    //累計金額
                    if (null != item.CUR_CUM_AMOUNT && item.CUR_CUM_AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[18].SetCellValue(double.Parse(item.CUR_CUM_AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[18].SetCellValue("");
                    }
                    //累計占比 
                    if (null != item.CUR_CUM_RATIO && item.CUR_CUM_RATIO.ToString().Trim() != "")
                    {
                        row.Cells[19].SetCellValue(double.Parse((item.CUR_CUM_RATIO / 100).ToString()));
                    }
                    else
                    {
                        row.Cells[19].SetCellValue("");
                    }
                    //會計科目名稱
                    logger.Debug("SUBJECT_NAM=" + item.SUBJECT_NAME);
                    row.Cells[20].SetCellValue(item.SUBJECT_NAME);
                    logger.Debug("get Site Expense Form cell style rowid=" + idxRow);
                    if (idxRow == 47)
                    {
                        idxRow = idxRow + 4;
                    }
                    idxRow++;
                }
            }
            else
            {
                float page = 1;
                logger.Info("公司費用單頁數 =" + Convert.ToInt16(Math.Ceiling(float.Parse(EXPTableItem.Count.ToString()) / 47)));
                if (float.Parse(EXPTableItem.Count.ToString()) / 47 > 1)
                {
                    page = float.Parse(EXPTableItem.Count.ToString()) / 47;
                    sheet = (XSSFSheet)hssfworkbook.GetSheet("費用表_2頁");
                }
                else
                {
                    sheet = (XSSFSheet)hssfworkbook.GetSheet("費用表");
                }
                //2.填入表頭資料
                sheet.GetRow(0).Cells[4].SetCellValue("協成水電工程事業有限公司 公司費用單");
                sheet.GetRow(1).Cells[3].SetCellValue("公司營業費用");
                sheet.GetRow(0).Cells[20].SetCellValue("第 1 頁");
                sheet.GetRow(0).Cells[22].SetCellValue("共" + Convert.ToInt16(Math.Ceiling(page)) + "頁");
                logger.Debug("Table Head_2=" + sheet.GetRow(1).Cells[9].ToString());
                sheet.GetRow(1).Cells[9].SetCellValue(ExpTable.EXP_FORM_ID);//費用單編號
                logger.Debug("Table Head_3=" + sheet.GetRow(1).Cells[14].ToString());
                sheet.GetRow(1).Cells[14].SetCellValue(ExpTable.PAYEE);//請款人(受款人)
                logger.Debug("Table Head_4=" + sheet.GetRow(1).Cells[19].ToString());
                sheet.GetRow(1).Cells[18].SetCellValue(ExpTable.OCCURRED_YEAR.ToString() + "/" + ExpTable.OCCURRED_MONTH.ToString());//費用發生年月
                logger.Debug("Table Head_5=" + sheet.GetRow(1).Cells[22].ToString());
                sheet.GetRow(1).Cells[22].SetCellValue(Convert.ToDateTime(ExpTable.CREATE_DATE).ToShortDateString());//請款日期
                if (float.Parse(EXPTableItem.Count.ToString()) / 47 > 1)
                {
                    sheet.GetRow(48).Cells[4].SetCellValue("協成水電工程事業有限公司 公司費用單");
                    sheet.GetRow(48).Cells[20].SetCellValue("第 2 頁");
                    sheet.GetRow(48).Cells[22].SetCellValue("共" + Convert.ToInt16(Math.Ceiling(page)) + "頁");
                    sheet.GetRow(49).Cells[3].SetCellValue("公司營業費用");
                    sheet.GetRow(49).Cells[9].SetCellValue(ExpTable.EXP_FORM_ID);//費用單編號
                    sheet.GetRow(49).Cells[14].SetCellValue(ExpTable.PAYEE);//請款人(受款人)
                    sheet.GetRow(49).Cells[18].SetCellValue(ExpTable.OCCURRED_YEAR.ToString() + "/" + ExpTable.OCCURRED_MONTH.ToString());//費用發生年月
                    sheet.GetRow(49).Cells[22].SetCellValue(Convert.ToDateTime(ExpTable.CREATE_DATE).ToShortDateString());//請款日期
                }
                if (float.Parse(EXPTableItem.Count.ToString()) / 47 > 1)
                {
                    logger.Debug("Table Head_6=" + sheet.GetRow(86).Cells[1].ToString());
                    sheet.GetRow(86).Cells[1].SetCellValue(ExpTable.REMARK);//說明事項
                    logger.Debug("Table Head_7=" + sheet.GetRow(86).Cells[12].ToString());
                    if (Budget.TOTAL_BUDGET.ToString() != "")
                    {
                        sheet.GetRow(86).Cells[13].SetCellValue(double.Parse(Budget.TOTAL_BUDGET.ToString()));//總預算
                    }
                    logger.Debug("Table Head_8=" + sheet.GetRow(89).Cells[12].ToString());
                    if (EarlyExpAmt.AMOUNT.ToString() != "")
                    {
                        sheet.GetRow(89).Cells[13].SetCellValue(double.Parse(EarlyExpAmt.AMOUNT.ToString()));//前期累計
                    }
                    else
                    {
                        sheet.GetRow(89).Cells[13].SetCellValue("0");
                    }
                    logger.Debug("Table Head_9=" + sheet.GetRow(90).Cells[12].ToString());
                    sheet.GetRow(90).Cells[13].SetCellValue(double.Parse(ExpAmt.AMOUNT.ToString()));//本期金額
                    logger.Debug("Table Head_10=" + sheet.GetRow(91).Cells[12].ToString());                                                                              //工資預算
                    sheet.GetRow(91).Cells[13].CellFormula = "(N90+N91)";
                    logger.Debug("Table Head_11=" + sheet.GetRow(92).Cells[12].ToString());
                    sheet.GetRow(92).Cells[13].SetCellValue(Convert.ToDateTime(ExpTable.PAYMENT_DATE).ToShortDateString());//支付日期
                }
                else
                {
                    logger.Debug("Table Head_6=" + sheet.GetRow(40).Cells[1].ToString());
                    sheet.GetRow(40).Cells[1].SetCellValue(ExpTable.REMARK);//說明事項
                    logger.Debug("Table Head_7=" + sheet.GetRow(40).Cells[12].ToString());
                    if (Budget.TOTAL_BUDGET.ToString() != "")
                    {
                        sheet.GetRow(40).Cells[13].SetCellValue(double.Parse(Budget.TOTAL_BUDGET.ToString()));//總預算
                    }
                    logger.Debug("Table Head_8=" + sheet.GetRow(43).Cells[12].ToString());
                    if (EarlyExpAmt.AMOUNT.ToString() != "")
                    {
                        sheet.GetRow(43).Cells[13].SetCellValue(double.Parse(EarlyExpAmt.AMOUNT.ToString()));//前期累計
                    }
                    else
                    {
                        sheet.GetRow(43).Cells[13].SetCellValue("0");
                    }
                    logger.Debug("Table Head_9=" + sheet.GetRow(44).Cells[12].ToString());
                    sheet.GetRow(44).Cells[13].SetCellValue(double.Parse(ExpAmt.AMOUNT.ToString()));//本期金額
                    logger.Debug("Table Head_10=" + sheet.GetRow(45).Cells[12].ToString());                                                                              //工資預算
                    sheet.GetRow(45).Cells[13].CellFormula = "(N44+N45)";
                    logger.Debug("Table Head_11=" + sheet.GetRow(46).Cells[12].ToString());
                    sheet.GetRow(46).Cells[13].SetCellValue(Convert.ToDateTime(ExpTable.PAYMENT_DATE).ToShortDateString());//支付日期
                }
                //3.填入資料
                int idxRow = 4;
                foreach (ExpenseBudgetSummary item in EXPTableItem)
                {
                    IRow row = sheet.GetRow(idxRow);//.CreateRow(idxRow);
                    logger.Info("Row Id=" + idxRow);
                    //項次、品名/摘要、單位、預算(合約)金額、前期累計金額、本期金額、累計金額、累計金額比率、會計名稱
                    //項次
                    row.Cells[0].SetCellValue(item.NO);
                    //品名或摘要
                    logger.Debug("ITEM_REMARK=" + item.ITEM_REMARK);
                    row.Cells[1].SetCellValue(item.ITEM_REMARK);
                    //單位
                    row.Cells[7].SetCellValue(item.ITEM_UNIT);
                    //預算(合約)金額
                    if (null != item.BUDGET_AMOUNT && item.BUDGET_AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[9].SetCellValue(double.Parse(item.BUDGET_AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[9].SetCellValue("");
                    }
                    //前期累計金額
                    if (null != item.CUM_AMOUNT && item.CUM_AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[14].SetCellValue(double.Parse(item.CUM_AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[14].SetCellValue("");
                    }
                    //本期金額
                    if (null != item.AMOUNT && item.AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[16].SetCellValue(double.Parse(item.AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[16].SetCellValue("");
                    }
                    //累計金額
                    if (null != item.CUR_CUM_AMOUNT && item.CUR_CUM_AMOUNT.ToString().Trim() != "")
                    {
                        row.Cells[18].SetCellValue(double.Parse(item.CUR_CUM_AMOUNT.ToString()));
                    }
                    else
                    {
                        row.Cells[18].SetCellValue("");
                    }
                    
                    //累計占比 
                    if (null != item.CUR_CUM_RATIO && item.CUR_CUM_RATIO.ToString().Trim() != "")
                    {
                        row.Cells[19].SetCellValue(double.Parse((item.CUR_CUM_RATIO / 100).ToString()));
                    }
                    else
                    {
                        row.Cells[19].SetCellValue("");
                    }
                    //會計科目名稱
                    logger.Debug("SUBJECT_NAM=" + item.SUBJECT_NAME);
                    row.Cells[20].SetCellValue(item.SUBJECT_NAME);
                    logger.Debug("get Expense Form cell style rowid=" + idxRow);
                    if (idxRow == 47)
                    {
                        idxRow = idxRow + 4;
                    }
                    idxRow++;
                }
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            if (null != ExpTable.PROJECT_ID && ExpTable.PROJECT_ID != "")
            {
                fileLocation = outputPath + "\\" + ExpTable.PROJECT_ID + "\\" + ExpTable.PROJECT_ID + "_" + ExpTable.EXP_FORM_ID + "_費用表.xlsx";
            }
            else
            {
                fileLocation = outputPath + "\\" + ExpTable.PROJECT_ID + "\\" + ExpTable.EXP_FORM_ID + "_公司費用表.xlsx";
            }
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public ExpenseFormToExcel()
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

    #region 公司費用預算執行彙整下載表格格式處理區段
    public class ExpBudgetSummaryToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string summaryFile = ContextService.strUploadPath + "\\ExpBudgetSummary_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        XSSFCellStyle style = null;
        XSSFCellStyle styleNumber = null;
        string fileformat = "xlsx";

        //存放費用表資料
        PurchaseFormService service = new PurchaseFormService();
        ExpenseBudgetSummary Amt = null;
        ExpenseBudgetSummary ExpAmt = null;
        public string totalBudget = null;
        public string totalExpense = null;
        public string errorMessage = null;

        //建立公司費用預算執行彙整下載表格
        public string exportExcel(int year)
        {
            List<ExpenseBudgetSummary> ExpBudget = service.getExpBudgetSummaryByYear(year);
            List<ExpenseBudgetByMonth> BudgetByMonth = service.getExpBudgetOfMonthByYear(year);
            List<ExpensetFromOPByMonth> ExpenseByMonth = service.getExpensetOfMonthByYear(year);
            Amt = service.getTotalExpBudgetAmount(year);
            totalBudget = String.Format("{0:#,##0.#}", Amt.TOTAL_BUDGET);
            ExpAmt = service.getTotalOperationExpAmount(year);
            totalExpense = String.Format("{0:#,##0.#}", ExpAmt.TOTAL_OPERATION_EXP);
            //1.讀取費用表格檔案
            InitializeWorkbook(summaryFile);
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("彙整表");
            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(year);//公司費用預算年度
            sheet.GetRow(2).Cells[15].SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));//製表日期
            sheet.GetRow(0).Cells[16].SetCellValue(totalBudget);//總預算金額
            sheet.GetRow(1).Cells[10].SetCellValue(totalExpense);//總執行金額

            //3.填入資料
            int idxRow = 5;
            foreach (ExpenseBudgetSummary item in ExpBudget)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //項目、項目代碼、預算/費用、7月、8月、9月、10月、11月、12月、1月、2月、3月、4月、5月、6月、合計
                //項目
                logger.Debug("SUBJECT_NAM=" + item.SUBJECT_NAME);
                row.CreateCell(0).SetCellValue(item.SUBJECT_NAME);
                row.Cells[0].CellStyle = style;
                //項目代碼
                row.CreateCell(1).SetCellValue(item.SUBJECT_ID);
                row.Cells[1].CellStyle = style;
                //預算/實際
                row.CreateCell(2).SetCellValue(item.NO);
                ICell cel3 = row.CreateCell(3);//設公式
                cel3.CellFormula = "IF(ISBLANK(C" + (idxRow + 1) + "),\"\",IF(MOD(C" + (idxRow + 1) + ",2)=0,\"實際數\",\"預算數\"))";
                cel3.CellStyle = style;

                row.CreateCell(4);//7月
                if (null != item.JUL && item.JUL.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.JUL.ToString()));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.CreateCell(5);//8月
                if (null != item.AUG && item.AUG.ToString().Trim() != "")
                {
                    row.Cells[5].SetCellValue(double.Parse(item.AUG.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//9月
                if (null != item.SEP && item.SEP.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellValue(double.Parse(item.SEP.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.CreateCell(7);//10月
                if (null != item.OCT && item.OCT.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellValue(double.Parse(item.OCT.ToString()));
                }
                else
                {
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[7].CellStyle = styleNumber;
                row.CreateCell(8);//11月
                if (null != item.NOV && item.NOV.ToString().Trim() != "")
                {
                    row.Cells[8].SetCellValue(double.Parse(item.NOV.ToString()));
                }
                else
                {
                    row.Cells[8].SetCellValue("");
                }
                row.Cells[8].CellStyle = styleNumber;
                row.CreateCell(9);//12月
                if (null != item.DEC && item.DEC.ToString().Trim() != "")
                {
                    row.Cells[9].SetCellValue(double.Parse(item.DEC.ToString()));
                }
                else
                {
                    row.Cells[9].SetCellValue("");
                }
                row.Cells[9].CellStyle = styleNumber;
                row.CreateCell(10);//1月
                if (null != item.JAN && item.JAN.ToString().Trim() != "")
                {
                    row.Cells[10].SetCellValue(double.Parse(item.JAN.ToString()));
                }
                else
                {
                    row.Cells[10].SetCellValue("");
                }
                row.Cells[10].CellStyle = styleNumber;
                row.CreateCell(11);//2月
                if (null != item.FEB && item.FEB.ToString().Trim() != "")
                {
                    row.Cells[11].SetCellValue(double.Parse(item.FEB.ToString()));
                }
                else
                {
                    row.Cells[11].SetCellValue("");
                }
                row.Cells[11].CellStyle = styleNumber;
                row.CreateCell(12);//3月
                if (null != item.MAR && item.MAR.ToString().Trim() != "")
                {
                    row.Cells[12].SetCellValue(double.Parse(item.MAR.ToString()));
                }
                else
                {
                    row.Cells[12].SetCellValue("");
                }
                row.Cells[12].CellStyle = styleNumber;
                row.CreateCell(13);//4月
                if (null != item.APR && item.APR.ToString().Trim() != "")
                {
                    row.Cells[13].SetCellValue(double.Parse(item.APR.ToString()));
                }
                else
                {
                    row.Cells[13].SetCellValue("");
                }
                row.Cells[13].CellStyle = styleNumber;
                row.CreateCell(14);//5月
                if (null != item.MAY && item.MAY.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.MAY.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                row.Cells[14].CellStyle = styleNumber;
                row.CreateCell(15);//6月
                if (null != item.JUN && item.JUN.ToString().Trim() != "")
                {
                    row.Cells[15].SetCellValue(double.Parse(item.JUN.ToString()));
                }
                else
                {
                    row.Cells[15].SetCellValue("");
                }
                row.Cells[15].CellStyle = styleNumber;
                row.CreateCell(16);//合計
                if (null != item.HTOTAL && item.HTOTAL.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.HTOTAL.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                row.Cells[16].CellStyle = styleNumber;
                logger.Debug("get ExpBudget cell style rowid=" + idxRow);
                idxRow++;
            }
            foreach (ExpenseBudgetByMonth item in BudgetByMonth)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //預算/費用、7月、8月、9月、10月、11月、12月、1月、2月、3月、4月、5月、6月、合計
                row.CreateCell(0);
                row.Cells[0].SetCellValue("合計");
                row.Cells[0].CellStyle = style;
                row.CreateCell(1);
                row.Cells[1].SetCellValue("");
                row.Cells[1].CellStyle = style;
                row.CreateCell(2);
                row.Cells[2].SetCellValue("");
                row.Cells[2].CellStyle = style;
                row.CreateCell(3);//預算
                row.Cells[3].SetCellValue("預算數");
                row.Cells[3].CellStyle = style;
                row.CreateCell(4);//7月
                if (null != item.JUL && item.JUL.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.JUL.ToString()));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.CreateCell(5);//8月
                if (null != item.AUG && item.AUG.ToString().Trim() != "")
                {
                    row.Cells[5].SetCellValue(double.Parse(item.AUG.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//9月
                if (null != item.SEP && item.SEP.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellValue(double.Parse(item.SEP.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.CreateCell(7);//10月
                if (null != item.OCT && item.OCT.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellValue(double.Parse(item.OCT.ToString()));
                }
                else
                {
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[7].CellStyle = styleNumber;
                row.CreateCell(8);//11月
                if (null != item.NOV && item.NOV.ToString().Trim() != "")
                {
                    row.Cells[8].SetCellValue(double.Parse(item.NOV.ToString()));
                }
                else
                {
                    row.Cells[8].SetCellValue("");
                }
                row.Cells[8].CellStyle = styleNumber;
                row.CreateCell(9);//12月
                if (null != item.DEC && item.DEC.ToString().Trim() != "")
                {
                    row.Cells[9].SetCellValue(double.Parse(item.DEC.ToString()));
                }
                else
                {
                    row.Cells[9].SetCellValue("");
                }
                row.Cells[9].CellStyle = styleNumber;
                row.CreateCell(10);//1月
                if (null != item.JAN && item.JAN.ToString().Trim() != "")
                {
                    row.Cells[10].SetCellValue(double.Parse(item.JAN.ToString()));
                }
                else
                {
                    row.Cells[10].SetCellValue("");
                }
                row.Cells[10].CellStyle = styleNumber;
                row.CreateCell(11);//2月
                if (null != item.FEB && item.FEB.ToString().Trim() != "")
                {
                    row.Cells[11].SetCellValue(double.Parse(item.FEB.ToString()));
                }
                else
                {
                    row.Cells[11].SetCellValue("");
                }
                row.Cells[11].CellStyle = styleNumber;
                row.CreateCell(12);//3月
                if (null != item.MAR && item.MAR.ToString().Trim() != "")
                {
                    row.Cells[12].SetCellValue(double.Parse(item.MAR.ToString()));
                }
                else
                {
                    row.Cells[12].SetCellValue("");
                }
                row.Cells[12].CellStyle = styleNumber;
                row.CreateCell(13);//4月
                if (null != item.APR && item.APR.ToString().Trim() != "")
                {
                    row.Cells[13].SetCellValue(double.Parse(item.APR.ToString()));
                }
                else
                {
                    row.Cells[13].SetCellValue("");
                }
                row.Cells[13].CellStyle = styleNumber;
                row.CreateCell(14);//5月
                if (null != item.MAY && item.MAY.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.MAY.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                row.Cells[14].CellStyle = styleNumber;
                row.CreateCell(15);//6月
                if (null != item.JUN && item.JUN.ToString().Trim() != "")
                {
                    row.Cells[15].SetCellValue(double.Parse(item.JUN.ToString()));
                }
                else
                {
                    row.Cells[15].SetCellValue("");
                }
                row.Cells[15].CellStyle = styleNumber;
                row.CreateCell(16);//合計
                if (null != item.HTOTAL && item.HTOTAL.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.HTOTAL.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                row.Cells[16].CellStyle = styleNumber;
                logger.Debug("get BudgetByMonth cell style rowid=" + idxRow);
                idxRow++;
            }
            foreach (ExpensetFromOPByMonth item in ExpenseByMonth)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //預算/費用、7月、8月、9月、10月、11月、12月、1月、2月、3月、4月、5月、6月、合計
                row.CreateCell(0);
                row.Cells[0].SetCellValue("合計");
                row.Cells[0].CellStyle = style;
                row.CreateCell(1);
                row.Cells[1].SetCellValue("");
                row.Cells[1].CellStyle = style;
                row.CreateCell(2);
                row.Cells[2].SetCellValue("");
                row.Cells[2].CellStyle = style;
                row.CreateCell(3);//實際
                row.Cells[3].SetCellValue("實際數");
                row.Cells[3].CellStyle = style;
                row.CreateCell(4);//7月
                if (null != item.JUL && item.JUL.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.JUL.ToString()));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.CreateCell(5);//8月
                if (null != item.AUG && item.AUG.ToString().Trim() != "")
                {
                    row.Cells[5].SetCellValue(double.Parse(item.AUG.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//9月
                if (null != item.SEP && item.SEP.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellValue(double.Parse(item.SEP.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.CreateCell(7);//10月
                if (null != item.OCT && item.OCT.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellValue(double.Parse(item.OCT.ToString()));
                }
                else
                {
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[7].CellStyle = styleNumber;
                row.CreateCell(8);//11月
                if (null != item.NOV && item.NOV.ToString().Trim() != "")
                {
                    row.Cells[8].SetCellValue(double.Parse(item.NOV.ToString()));
                }
                else
                {
                    row.Cells[8].SetCellValue("");
                }
                row.Cells[8].CellStyle = styleNumber;
                row.CreateCell(9);//12月
                if (null != item.DEC && item.DEC.ToString().Trim() != "")
                {
                    row.Cells[9].SetCellValue(double.Parse(item.DEC.ToString()));
                }
                else
                {
                    row.Cells[9].SetCellValue("");
                }
                row.Cells[9].CellStyle = styleNumber;
                row.CreateCell(10);//1月
                if (null != item.JAN && item.JAN.ToString().Trim() != "")
                {
                    row.Cells[10].SetCellValue(double.Parse(item.JAN.ToString()));
                }
                else
                {
                    row.Cells[10].SetCellValue("");
                }
                row.Cells[10].CellStyle = styleNumber;
                row.CreateCell(11);//2月
                if (null != item.FEB && item.FEB.ToString().Trim() != "")
                {
                    row.Cells[11].SetCellValue(double.Parse(item.FEB.ToString()));
                }
                else
                {
                    row.Cells[11].SetCellValue("");
                }
                row.Cells[11].CellStyle = styleNumber;
                row.CreateCell(12);//3月
                if (null != item.MAR && item.MAR.ToString().Trim() != "")
                {
                    row.Cells[12].SetCellValue(double.Parse(item.MAR.ToString()));
                }
                else
                {
                    row.Cells[12].SetCellValue("");
                }
                row.Cells[12].CellStyle = styleNumber;
                row.CreateCell(13);//4月
                if (null != item.APR && item.APR.ToString().Trim() != "")
                {
                    row.Cells[13].SetCellValue(double.Parse(item.APR.ToString()));
                }
                else
                {
                    row.Cells[13].SetCellValue("");
                }
                row.Cells[13].CellStyle = styleNumber;
                row.CreateCell(14);//5月
                if (null != item.MAY && item.MAY.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.MAY.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                row.Cells[14].CellStyle = styleNumber;
                row.CreateCell(15);//6月
                if (null != item.JUN && item.JUN.ToString().Trim() != "")
                {
                    row.Cells[15].SetCellValue(double.Parse(item.JUN.ToString()));
                }
                else
                {
                    row.Cells[15].SetCellValue("");
                }
                row.Cells[15].CellStyle = styleNumber;
                row.CreateCell(16);//合計
                if (null != item.HTOTAL && item.HTOTAL.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.HTOTAL.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                row.Cells[16].CellStyle = styleNumber;
                logger.Debug("get ExpenseByMonth cell style rowid=" + idxRow);
                idxRow++;
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            fileLocation = outputPath + "\\" + year + "_公司費用預算執行彙整表.xlsx";
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public ExpBudgetSummaryToExcel()
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

    #region 物料採購單與領料單下載表格格式處理區段
    public class MaterialFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string orderFile = ContextService.strUploadPath + "\\MaterialOrder_form.xlsx";
        string deliveryFile = ContextService.strUploadPath + "\\MaterialDelivery_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        XSSFCellStyle style = null;
        XSSFCellStyle styleNumber = null;

        string fileformat = "xlsx";
        //存放物料單資料
        public TND_PROJECT project = null;
        public PLAN_PURCHASE_REQUISITION tablePR = null;
        public List<PurchaseRequisition> orderItem = null;
        public List<PurchaseRequisition> deliveryItem = null;
        public string errorMessage = null;


        //建立物料單下載表格
        public string exportExcel(TND_PROJECT project, PLAN_PURCHASE_REQUISITION tablePR, List<PurchaseRequisition> orderItem, List<PurchaseRequisition> deliveryItem, bool isOrder, bool isDO)
        {
            //1.讀取預算表格檔案
            if (isOrder)
            {
                InitializeWorkbook(orderFile);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("採購單");
            }
            else
            {
                InitializeWorkbook(deliveryFile);
                sheet = (XSSFSheet)hssfworkbook.GetSheet("領料單");
            }
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
            //2.填入表頭資料
            if (isOrder)
            {
                logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
                sheet.GetRow(1).Cells[2].SetCellValue(project.PROJECT_NAME);//專案名稱
                logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
                sheet.GetRow(2).Cells[2].SetCellValue(tablePR.SUPPLIER_ID);//供應商名稱
                logger.Debug("Table Head_3=" + sheet.GetRow(3).Cells[0].ToString());
                sheet.GetRow(3).Cells[2].SetCellValue(tablePR.PR_ID);//採購單號
                logger.Debug("Table Head_4=" + sheet.GetRow(5).Cells[0].ToString());
                sheet.GetRow(5).Cells[2].SetCellValue(tablePR.RECIPIENT);//收件人
                logger.Debug("Table Head_5=" + sheet.GetRow(6).Cells[0].ToString());
                sheet.GetRow(6).Cells[2].SetCellValue(tablePR.LOCATION);//送貨地址
                logger.Debug("Table Head_6=" + sheet.GetRow(7).Cells[0].ToString());
                sheet.GetRow(7).Cells[2].SetCellValue(tablePR.REMARK);//注意事項
            }
            else
            {
                logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
                sheet.GetRow(1).Cells[2].SetCellValue(project.PROJECT_NAME);//專案名稱
                logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
                sheet.GetRow(2).Cells[2].SetCellValue(Convert.ToDateTime(tablePR.CREATE_DATE).ToString("yyyy/MM/dd"));//領收日期
                logger.Debug("Table Head_3=" + sheet.GetRow(3).Cells[0].ToString());
                sheet.GetRow(3).Cells[2].SetCellValue(tablePR.PR_ID);//領料單號
                logger.Debug("Table Head_4=" + sheet.GetRow(4).Cells[0].ToString());
                sheet.GetRow(4).Cells[2].SetCellValue(tablePR.RECIPIENT);//領收人所屬單位/公司
                if (isDO)
                {
                    //var title = sheet.CreateRow(0);
                    //title.CreateCell(0);
                    //title.Cells[0].SetCellValue("領料單");
                    //title.Cells[0].CellStyle.Alignment = HorizontalAlignment.Center; //水平置中;
                    logger.Debug("Table Head_0=" + sheet.GetRow(0).Cells[2].ToString());
                    sheet.GetRow(0).Cells[3].SetCellValue("領料單");
                    logger.Debug("Table Head_5=" + sheet.GetRow(6).Cells[4].ToString());
                    sheet.GetRow(6).Cells[5].SetCellValue("領收數量");
                }
                else
                {
                    logger.Debug("Table Head_0=" + sheet.GetRow(0).Cells[2].ToString());
                    sheet.GetRow(0).Cells[3].SetCellValue("領料單(無標單品項)");
                    logger.Debug("Table Head_5=" + sheet.GetRow(6).Cells[4].ToString());
                    sheet.GetRow(6).Cells[5].SetCellValue("驗收數量");
                }
            }
            //3.填入資料
            if (isOrder)
            {
                int idxRow = 10;
                foreach (PurchaseRequisition item in orderItem)
                {
                    IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                    logger.Info("Row Id=" + idxRow);
                    //No.、項次、、項目說明、單位、合約數量、採購數量、需求日期、備註
                    row.CreateCell(0);
                    row.Cells[0].SetCellValue(item.NO);//No.
                    row.Cells[0].CellStyle = style;
                    row.CreateCell(1);
                    row.Cells[1].SetCellValue(item.ITEM_ID);//項次
                    row.Cells[1].CellStyle = style;
                    logger.Debug("ITEM DESC=" + item.ITEM_DESC);
                    row.CreateCell(2);
                    row.Cells[2].SetCellValue(item.ITEM_DESC);//項目說明
                    row.Cells[2].CellStyle = style;
                    row.CreateCell(3);
                    row.Cells[3].SetCellValue(item.ITEM_UNIT);// 單位
                    row.Cells[3].CellStyle = style;
                    row.CreateCell(4);
                    if (null != item.MAP_QTY && item.MAP_QTY.ToString().Trim() != "")
                    {
                        row.Cells[4].SetCellValue(double.Parse(item.MAP_QTY.ToString())); //合約數量
                    }
                    row.Cells[4].CellStyle = style;
                    row.CreateCell(5);
                    if (null != item.ORDER_QTY && item.ORDER_QTY.ToString().Trim() != "")
                    {
                        row.Cells[5].SetCellValue(double.Parse(item.ORDER_QTY.ToString())); //採購數量
                    }
                    row.Cells[5].CellStyle = style;
                    row.CreateCell(6);
                    if (null != item.NEED_DATE && item.NEED_DATE.ToString().Trim() != "")
                    {
                        row.Cells[6].SetCellValue(Convert.ToDateTime(item.NEED_DATE).ToString("yyyy/MM/dd")); //需求日期
                    }
                    row.Cells[6].CellStyle = style;
                    row.CreateCell(7);
                    row.Cells[7].SetCellValue(item.ITEM_REMARK);//備註
                    row.Cells[7].CellStyle = style;
                    logger.Debug("get Material Order Form cell style rowid=" + idxRow);
                    idxRow++;
                }
                idxRow++;
                IRow nextRow = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Next Row Id=" + idxRow);
                nextRow.CreateCell(0);
                nextRow.Cells[0].SetCellValue("特殊需求 :");
                idxRow++;
                IRow lastRow = sheet.CreateRow(idxRow);
                lastRow.CreateCell(0);
                lastRow.Cells[0].SetCellValue(tablePR.MESSAGE);//特殊需求
            }
            else if (isDO)
            {
                int idxRow = 7;
                foreach (PurchaseRequisition item in deliveryItem)
                {
                    IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                    logger.Info("Row Id=" + idxRow);
                    //No.、項次、、項目說明、單位、主系統、領收數量
                    row.CreateCell(0);
                    row.Cells[0].SetCellValue(item.NO);//No.
                    row.Cells[0].CellStyle = style;
                    row.CreateCell(1);
                    row.Cells[1].SetCellValue(item.ITEM_ID);//項次
                    row.Cells[1].CellStyle = style;
                    logger.Debug("ITEM DESC=" + item.ITEM_DESC);
                    row.CreateCell(2);
                    row.Cells[2].SetCellValue(item.ITEM_DESC);//項目說明
                    row.Cells[2].CellStyle = style;
                    row.CreateCell(3);
                    row.Cells[3].SetCellValue(item.ITEM_UNIT);// 單位
                    row.Cells[3].CellStyle = style;
                    row.CreateCell(4);
                    row.Cells[4].SetCellValue(item.SYSTEM_MAIN);// 主系統
                    row.Cells[4].CellStyle = style;
                    row.CreateCell(5);
                    if (null != item.DELIVERY_QTY && item.DELIVERY_QTY.ToString().Trim() != "")
                    {
                        row.Cells[5].SetCellValue(double.Parse(item.DELIVERY_QTY.ToString())); //領收數量
                    }
                    row.Cells[5].CellStyle = style;
                    logger.Debug("get Material DO Form cell style rowid=" + idxRow);
                    idxRow++;
                }
                idxRow++;
                IRow newRow = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("New Row Id=" + idxRow);
                newRow.CreateCell(0);
                newRow.Cells[0].SetCellValue("領料說明 :");
                idxRow++;
                IRow nextRow = sheet.CreateRow(idxRow);
                logger.Info("Next Row Id=" + idxRow);
                nextRow.CreateCell(0);
                nextRow.Cells[0].SetCellValue(tablePR.CAUTION);//領料說明
                idxRow = idxRow + 2;
                IRow lastRow = sheet.CreateRow(idxRow);
                lastRow.CreateCell(0);
                lastRow.CreateCell(1);
                lastRow.CreateCell(2);
                lastRow.CreateCell(3);
                lastRow.CreateCell(4);
                lastRow.Cells[4].SetCellValue("領收人簽名 :");//領收人簽名
            }
            else
            {
                int idxRow = 7;
                foreach (PurchaseRequisition item in orderItem)
                {
                    IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                    logger.Info("Row Id=" + idxRow);
                    //No.、項次、、項目說明、單位、主系統、領收數量
                    row.CreateCell(0);
                    row.Cells[0].SetCellValue(item.NO);//No.
                    row.Cells[0].CellStyle = style;
                    row.CreateCell(1);
                    row.Cells[1].SetCellValue(item.ITEM_ID);//項次
                    row.Cells[1].CellStyle = style;
                    logger.Debug("ITEM DESC=" + item.ITEM_DESC);
                    row.CreateCell(2);
                    row.Cells[2].SetCellValue(item.ITEM_DESC);//項目說明
                    row.Cells[2].CellStyle = style;
                    row.CreateCell(3);
                    row.Cells[3].SetCellValue(item.ITEM_UNIT);// 單位
                    row.Cells[3].CellStyle = style;
                    row.CreateCell(4);
                    row.Cells[4].SetCellValue(item.SYSTEM_MAIN);// 主系統
                    row.Cells[4].CellStyle = style;
                    row.CreateCell(5);
                    if (null != item.RECEIPT_QTY && item.RECEIPT_QTY.ToString().Trim() != "")
                    {
                        row.Cells[5].SetCellValue(double.Parse(item.RECEIPT_QTY.ToString())); //驗收數量
                    }
                    row.Cells[5].CellStyle = style;
                    logger.Debug("get Material DF Form cell style rowid=" + idxRow);
                    idxRow++;
                }
                idxRow++;
                IRow newRow = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("New Row Id=" + idxRow);
                newRow.CreateCell(0);
                newRow.Cells[0].SetCellValue("領料說明 :");
                idxRow++;
                IRow nextRow = sheet.CreateRow(idxRow);
                logger.Info("Next Row Id=" + idxRow);
                nextRow.CreateCell(0);
                nextRow.Cells[0].SetCellValue(tablePR.CAUTION);//領料說明
                idxRow = idxRow + 2;
                IRow lastRow = sheet.CreateRow(idxRow);
                lastRow.CreateCell(0);
                lastRow.CreateCell(1);
                lastRow.CreateCell(2);
                lastRow.CreateCell(3);
                lastRow.CreateCell(4);
                lastRow.Cells[4].SetCellValue("領收人簽名 :");//領收人簽名
            }
            //lastRow.Cells[0].CellStyle = style;
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            if (isOrder)
            {
                fileLocation = outputPath + "\\" + project.PROJECT_ID + "\\" + tablePR.PR_ID + "_採購單.xlsx";
            }
            else
            {
                fileLocation = outputPath + "\\" + project.PROJECT_ID + "\\" + tablePR.PR_ID + "_領料單.xlsx";
            }
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public MaterialFormToExcel()
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
    #region 工地費用預算下載表格格式處理區段
    public class SiteBudgetFormToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string budgetFile = ContextService.strUploadPath + "\\site_budget_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放工地費用預算資料
        PurchaseFormService service = new PurchaseFormService();
        public List<FIN_SUBJECT> subjects = null;
        public string errorMessage = null;
        public TND_PROJECT project = null;
        string projId = null;
        string budgetYear = null;
        string sequenceYear = null;
        //建立工地費用預算下載表格
        public string exportExcel(TND_PROJECT project)
        {
            List<FIN_SUBJECT> subjects = service.getSubjectOfExpense4Site();
            //1.讀取工地費用預算表格檔案
            InitializeWorkbook(budgetFile);
            for (int i = 1; i < 6; i++)
            {
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工地費用預算_第" + i + "年度");

                //2.填入表頭資料
                logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
                sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
                logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
                sheet.GetRow(2).Cells[1].SetCellValue(project.PROJECT_NAME);//專案名稱
                sheet.GetRow(4).Cells[13].SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));//製表日期
                                                                                            //3.填入資料
                int idxRow = 7;
                foreach (FIN_SUBJECT item in subjects)
                {
                    IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                    logger.Info("Row Id=" + idxRow);
                    //項目、項目代碼
                    //項目
                    row.CreateCell(0).SetCellValue(item.SUBJECT_NAME);
                    //項目代碼
                    row.CreateCell(1).SetCellValue(item.FIN_SUBJECT_ID);
                    row.CreateCell(2).SetCellValue("");
                    row.CreateCell(3).SetCellValue("");
                    row.CreateCell(4).SetCellValue("");
                    row.CreateCell(5).SetCellValue("");
                    row.CreateCell(6).SetCellValue("");
                    row.CreateCell(7).SetCellValue("");
                    row.CreateCell(8).SetCellValue("");
                    row.CreateCell(9).SetCellValue("");
                    row.CreateCell(10).SetCellValue("");
                    row.CreateCell(11).SetCellValue("");
                    row.CreateCell(12).SetCellValue("");
                    row.CreateCell(13).SetCellValue("");
                    foreach (ICell c in row.Cells)
                    {
                        c.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                    }
                    ICell cel14 = row.CreateCell(14);
                    cel14.CellFormula = "C" + (idxRow + 1) + "+D" + (idxRow + 1) + "+E" + (idxRow + 1) + "+F" + (idxRow + 1) + "+G" + (idxRow + 1) + "+H" + (idxRow + 1)
                    + "+I" + (idxRow + 1) + "+J" + (idxRow + 1) + "+K" + (idxRow + 1) + "+L" + (idxRow + 1) + "+M" + (idxRow + 1) + "+N" + (idxRow + 1);
                    cel14.CellStyle = ExcelStyle.getNumberStyle(hssfworkbook);
                    logger.Debug("getSubject cell style rowid=" + idxRow);
                    idxRow++;
                }
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            fileLocation = outputPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "_工地費用預算.xlsx";
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public SiteBudgetFormToExcel()
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

        #region 工地費用預算資料轉換 
        /**
         * 取得工地費用預算Sheet資料
         * */
        #region 第1年度
        public List<PLAN_SITE_BUDGET> ConvertDataForSiteBudget1(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + "工地費用預算_第1年度");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工地費用預算_第1年度");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + "工地費用預算_第1年度");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工地費用預算_第1年度");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有:工地費用預算_第1年度(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工地費用預算_第1年度]資料");
            }
            return ConverData2SiteBudget();
        }
        #endregion
        #region 第2年度
        public List<PLAN_SITE_BUDGET> ConvertDataForSiteBudget2(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + "工地費用預算_第2年度");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工地費用預算_第2年度");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + "工地費用預算_第2年度");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工地費用預算_第2年度");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有:工地費用預算_第2度(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工地費用預算_第2年度]資料");
            }
            return ConverData2SiteBudget();
        }
        #endregion
        #region 第3年度
        public List<PLAN_SITE_BUDGET> ConvertDataForSiteBudget3(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + "工地費用預算_第3年度");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工地費用預算_第3年度");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + "工地費用預算_第3年度");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工地費用預算_第3年度");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有:工地費用預算_第3年度(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工地費用預算_第3年度]資料");
            }
            return ConverData2SiteBudget();
        }
        #endregion
        #region 第4年度
        public List<PLAN_SITE_BUDGET> ConvertDataForSiteBudget4(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + "工地費用預算_第4年度");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工地費用預算_第4年度");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + "工地費用預算_第4年度");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工地費用預算_第4年度");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有:工地費用預算_第4年度(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工地費用預算_第4年度]資料");
            }
            return ConverData2SiteBudget();
        }
        #endregion
        #region 第5年度
        public List<PLAN_SITE_BUDGET> ConvertDataForSiteBudget5(string projectId)
        {
            projId = projectId;
            //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟預算Sheet
            if (fileformat == "xls")
            {
                logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + "工地費用預算_第5年度");
                sheet = (HSSFSheet)hssfworkbook.GetSheet("工地費用預算_第5年度");
            }
            else
            {
                logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + "工地費用預算_第5年度");
                sheet = (XSSFSheet)hssfworkbook.GetSheet("工地費用預算_第5年度");
            }
            if (null == sheet)
            {
                logger.Error("檔案內沒有:工地費用預算_第5年度(Sheet)! filename=" + fileformat);
                throw new Exception("檔案內沒有[工地費用預算_第5年度]資料");
            }
            return ConverData2SiteBudget();
        }
        #endregion
        /**
         * 轉換工地費用預算資料檔:工地費用預算
         * */
        protected List<PLAN_SITE_BUDGET> ConverData2SiteBudget()
        {
            IRow row = null;
            List<PLAN_SITE_BUDGET> lstSiteBudget = new List<PLAN_SITE_BUDGET>();
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
            //2.逐行讀取資料
            int iRowIndex = 0; //0 表 Row 1

            //2.1  忽略不要的行數..(表頭)
            budgetYear = sheet.GetRow(4).Cells[1].ToString();
            sequenceYear = sheet.GetRow(3).Cells[1].ToString();
            while (iRowIndex < (6))
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
                //0.項目代碼 2.1月金額 3.2月金額 4.3月金額 5.4月金額 6.5月金額 7.6月金額 8.7月金額 9.8月金額 10.9月金額 11.10月金額 12.11月金額 13.12月金額
                if (row.Cells[0].ToString().ToUpper() != "END")
                {
                    List<PLAN_SITE_BUDGET> lst = convertRow2SiteBudget(row, iRowIndex);
                    foreach (PLAN_SITE_BUDGET it in lst)
                    {
                        lstSiteBudget.Add(it);
                    }
                }
                else
                {
                    logErrorMessage("Step1 ;取得工地費用預算資料:" + subjects.Count + "筆");
                    logger.Info("Finish convert Job : count=" + subjects.Count);
                    return lstSiteBudget;
                }
                iRowIndex++;
            }
            logger.Info("Plan_Site_Budget Count:" + iRowIndex);
            return lstSiteBudget;
        }
        /**
         * 將Excel Row 轉換成為對應的資料物件
         * */
        private List<PLAN_SITE_BUDGET> convertRow2SiteBudget(IRow row, int excelrow)
        {
            List<PLAN_SITE_BUDGET> lst = new List<PLAN_SITE_BUDGET>();
            String strItemId = row.Cells[1].ToString();
            if (null != strItemId && strItemId != "")
            {
                for (int i = 0; i < 12; i++)
                {
                    PLAN_SITE_BUDGET item = new PLAN_SITE_BUDGET();
                    item.PROJECT_ID = projId;
                    item.BUDGET_YEAR = int.Parse(budgetYear);
                    item.BUDGET_MONTH = i + 1;
                    item.SUBJECT_ID = row.Cells[1].ToString();
                    item.CREATE_DATE = System.DateTime.Now;
                    item.YEAR_SEQUENCE = sequenceYear;
                    if (null != row.Cells[i + 2].ToString().Trim() || row.Cells[i + 2].ToString().Trim() != "")
                    {
                        try
                        {
                            decimal dAmt = decimal.Parse(row.Cells[i + 2].ToString());
                            logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[i + 2].ToString());
                            item.AMOUNT = dAmt;
                        }
                        catch (Exception e)
                        {
                            logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[i+2].value=" + row.Cells[i + 2].ToString());
                            logger.Error(e);
                        }
                    }
                    lst.Add(item);
                }
            }
            return lst;
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

    #region 工地費用預算執行彙整下載表格格式處理區段
    public class SiteExpSummaryToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string summaryFile = ContextService.strUploadPath + "\\SiteExpSummary_form.xlsx";
        string outputPath = ContextService.strUploadPath;
        IWorkbook hssfworkbook;
        ISheet sheet = null;
        XSSFCellStyle style = null;
        XSSFCellStyle styleNumber = null;
        string fileformat = "xlsx";

        //存放費用表資料
        PurchaseFormService service = new PurchaseFormService();
        ExpenseBudgetSummary Amt = null;
        ExpenseBudgetSummary ExpAmt = null;
        public string totalBudget = null;
        public string totalExpense = null;
        public string errorMessage = null;

        //建立工地費用預算執行彙整下載表格
        public string exportExcel(string projectid, int sequence, int targetYear, int targetMonth, bool isCum)
        {
            List<ExpenseBudgetSummary> ExpBudget = service.getSiteExpBudgetSummaryBySeqYear(projectid, sequence, targetYear, targetMonth, isCum);
            List<ExpenseBudgetByMonth> BudgetByMonth = service.getSiteExpBudgetOfMonth(projectid, sequence, targetYear, targetMonth, isCum);
            List<ExpensetFromOPByMonth> ExpenseByMonth = service.getSiteExpensetOfMonth(projectid, targetYear, targetMonth, isCum);
            Amt = service.getSiteBudgetAmountById(projectid);
            totalBudget = String.Format("{0:#,##0.#}", Amt.TOTAL_BUDGET);
            ExpAmt = service.getTotalSiteExpAmountById(projectid, targetYear, targetMonth, isCum);
            totalExpense = String.Format("{0:#,##0.#}", ExpAmt.CUM_YEAR_AMOUNT);
            string date = targetYear.ToString() + "/" + targetMonth.ToString();
            //1.讀取費用表格檔案
            InitializeWorkbook(summaryFile);
            style = ExcelStyle.getContentStyle(hssfworkbook);
            styleNumber = ExcelStyle.getNumberStyle(hssfworkbook);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("彙整表");
            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            if (isCum == true)
            {
                sheet.GetRow(1).Cells[0].SetCellValue("條件為累計至 :");//工地費用查詢條件表頭文字
                sheet.GetRow(1).Cells[1].SetCellValue(date);//工地費用累計查詢年月
            }
            else
            {
                sheet.GetRow(1).Cells[0].SetCellValue("查詢年度 :");//工地費用查詢條件表頭文字
                sheet.GetRow(1).Cells[1].SetCellValue(targetYear);//工地費用查詢年度
            }
            sheet.GetRow(2).Cells[15].SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));//製表日期
            sheet.GetRow(0).Cells[16].SetCellValue(totalBudget);//總預算金額
            sheet.GetRow(1).Cells[10].SetCellValue(totalExpense);//執行金額
            if (isCum == true)
            {
                sheet.GetRow(3).Cells[16].SetCellValue("累計");//更改合計欄位為累計
            }
            //3.填入資料
            int idxRow = 5;
            foreach (ExpenseBudgetSummary item in ExpBudget)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //項目、項目代碼、預算/費用、1月、2月、3月、4月、5月、6月、7月、8月、9月、10月、11月、12月、合計
                //項目
                logger.Debug("SUBJECT_NAM=" + item.SUBJECT_NAME);
                row.CreateCell(0).SetCellValue(item.SUBJECT_NAME);
                row.Cells[0].CellStyle = style;
                //項目代碼
                row.CreateCell(1).SetCellValue(item.SUBJECT_ID);
                row.Cells[1].CellStyle = style;
                //預算/實際
                row.CreateCell(2).SetCellValue(item.NO);
                ICell cel3 = row.CreateCell(3);//設公式
                cel3.CellFormula = "IF(ISBLANK(C" + (idxRow + 1) + "),\"\",IF(MOD(C" + (idxRow + 1) + ",2)=0,\"實際數\",\"預算數\"))";
                cel3.CellStyle = style;

                row.CreateCell(4);//1月
                if (null != item.JAN && item.JAN.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.JAN.ToString()));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.CreateCell(5);//2月
                if (null != item.FEB && item.FEB.ToString().Trim() != "")
                {
                    row.Cells[5].SetCellValue(double.Parse(item.FEB.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//3月
                if (null != item.MAR && item.MAR.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellValue(double.Parse(item.MAR.ToString()));
                }
                else
                {
                    row.Cells[6].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.CreateCell(7);//4月
                if (null != item.APR && item.APR.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellValue(double.Parse(item.APR.ToString()));
                }
                else
                {
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[7].CellStyle = styleNumber;
                row.CreateCell(8);//5月
                if (null != item.MAY && item.MAY.ToString().Trim() != "")
                {
                    row.Cells[8].SetCellValue(double.Parse(item.MAY.ToString()));
                }
                else
                {
                    row.Cells[8].SetCellValue("");
                }
                row.Cells[8].CellStyle = styleNumber;
                row.CreateCell(9);//6月
                if (null != item.JUN && item.JUN.ToString().Trim() != "")
                {
                    row.Cells[9].SetCellValue(double.Parse(item.JUN.ToString()));
                }
                else
                {
                    row.Cells[9].SetCellValue("");
                }
                row.Cells[9].CellStyle = styleNumber;
                row.CreateCell(10);//7月
                if (null != item.JUL && item.JUL.ToString().Trim() != "")
                {
                    row.Cells[10].SetCellValue(double.Parse(item.JUL.ToString()));
                }
                else
                {
                    row.Cells[10].SetCellValue("");
                }
                row.Cells[10].CellStyle = styleNumber;
                row.CreateCell(11);//8月
                if (null != item.AUG && item.AUG.ToString().Trim() != "")
                {
                    row.Cells[11].SetCellValue(double.Parse(item.AUG.ToString()));
                }
                else
                {
                    row.Cells[11].SetCellValue("");
                }
                row.Cells[11].CellStyle = styleNumber;
                row.CreateCell(12);//9月
                if (null != item.SEP && item.SEP.ToString().Trim() != "")
                {
                    row.Cells[12].SetCellValue(double.Parse(item.SEP.ToString()));
                }
                else
                {
                    row.Cells[12].SetCellValue("");
                }
                row.Cells[12].CellStyle = styleNumber;
                row.CreateCell(13);//10月
                if (null != item.OCT && item.OCT.ToString().Trim() != "")
                {
                    row.Cells[13].SetCellValue(double.Parse(item.OCT.ToString()));
                }
                else
                {
                    row.Cells[13].SetCellValue("");
                }
                row.Cells[13].CellStyle = styleNumber;
                row.CreateCell(14);//11月
                if (null != item.NOV && item.NOV.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.NOV.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                row.Cells[14].CellStyle = styleNumber;
                row.CreateCell(15);//12月
                if (null != item.DEC && item.DEC.ToString().Trim() != "")
                {
                    row.Cells[15].SetCellValue(double.Parse(item.DEC.ToString()));
                }
                else
                {
                    row.Cells[15].SetCellValue("");
                }
                row.Cells[15].CellStyle = styleNumber;
                
                row.CreateCell(16);//合計
                if (null != item.HTOTAL && item.HTOTAL.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.HTOTAL.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                row.Cells[16].CellStyle = styleNumber;
                logger.Debug("get Site ExpBudget cell style rowid=" + idxRow);
                idxRow++;
            }
            foreach (ExpenseBudgetByMonth item in BudgetByMonth)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //預算/費用、1月、2月、3月、4月、5月、6月、7月、8月、9月、10月、11月、12月、合計
                row.CreateCell(0);
                row.Cells[0].SetCellValue("合計");
                row.Cells[0].CellStyle = style;
                row.CreateCell(1);
                row.Cells[1].SetCellValue("");
                row.Cells[1].CellStyle = style;
                row.CreateCell(2);
                row.Cells[2].SetCellValue("");
                row.Cells[2].CellStyle = style;
                row.CreateCell(3);//預算
                row.Cells[3].SetCellValue("預算數");
                row.Cells[3].CellStyle = style;
                row.CreateCell(4);//1月
                if (null != item.JAN && item.JAN.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.JAN.ToString()));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.CreateCell(5);//2月
                if (null != item.FEB && item.FEB.ToString().Trim() != "")
                {
                    row.Cells[5].SetCellValue(double.Parse(item.FEB.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//3月
                if (null != item.MAR && item.MAR.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellValue(double.Parse(item.MAR.ToString()));
                }
                else
                {
                    row.Cells[6].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.CreateCell(7);//4月
                if (null != item.APR && item.APR.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellValue(double.Parse(item.APR.ToString()));
                }
                else
                {
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[7].CellStyle = styleNumber;
                row.CreateCell(8);//5月
                if (null != item.MAY && item.MAY.ToString().Trim() != "")
                {
                    row.Cells[8].SetCellValue(double.Parse(item.MAY.ToString()));
                }
                else
                {
                    row.Cells[8].SetCellValue("");
                }
                row.Cells[8].CellStyle = styleNumber;
                row.CreateCell(9);//6月
                if (null != item.JUN && item.JUN.ToString().Trim() != "")
                {
                    row.Cells[9].SetCellValue(double.Parse(item.JUN.ToString()));
                }
                else
                {
                    row.Cells[9].SetCellValue("");
                }
                row.Cells[9].CellStyle = styleNumber;
                row.CreateCell(10);//7月
                if (null != item.JUL && item.JUL.ToString().Trim() != "")
                {
                    row.Cells[10].SetCellValue(double.Parse(item.JUL.ToString()));
                }
                else
                {
                    row.Cells[10].SetCellValue("");
                }
                row.Cells[10].CellStyle = styleNumber;
                row.CreateCell(11);//8月
                if (null != item.AUG && item.AUG.ToString().Trim() != "")
                {
                    row.Cells[11].SetCellValue(double.Parse(item.AUG.ToString()));
                }
                else
                {
                    row.Cells[11].SetCellValue("");
                }
                row.Cells[11].CellStyle = styleNumber;
                row.CreateCell(12);//9月
                if (null != item.SEP && item.SEP.ToString().Trim() != "")
                {
                    row.Cells[12].SetCellValue(double.Parse(item.SEP.ToString()));
                }
                else
                {
                    row.Cells[12].SetCellValue("");
                }
                row.Cells[12].CellStyle = styleNumber;
                row.CreateCell(13);//10月
                if (null != item.OCT && item.OCT.ToString().Trim() != "")
                {
                    row.Cells[13].SetCellValue(double.Parse(item.OCT.ToString()));
                }
                else
                {
                    row.Cells[13].SetCellValue("");
                }
                row.Cells[13].CellStyle = styleNumber;
                row.CreateCell(14);//11月
                if (null != item.NOV && item.NOV.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.NOV.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                row.Cells[14].CellStyle = styleNumber;
                row.CreateCell(15);//12月
                if (null != item.DEC && item.DEC.ToString().Trim() != "")
                {
                    row.Cells[15].SetCellValue(double.Parse(item.DEC.ToString()));
                }
                else
                {
                    row.Cells[15].SetCellValue("");
                }
                row.Cells[15].CellStyle = styleNumber;
                row.CreateCell(16);//合計
                if (null != item.HTOTAL && item.HTOTAL.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.HTOTAL.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                row.Cells[16].CellStyle = styleNumber;
                logger.Debug("get Site BudgetByMonth cell style rowid=" + idxRow);
                idxRow++;
            }
            foreach (ExpensetFromOPByMonth item in ExpenseByMonth)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                logger.Info("Row Id=" + idxRow);
                //預算/費用、1月、2月、3月、4月、5月、6月、7月、8月、9月、10月、11月、12月、合計
                row.CreateCell(0);
                row.Cells[0].SetCellValue("合計");
                row.Cells[0].CellStyle = style;
                row.CreateCell(1);
                row.Cells[1].SetCellValue("");
                row.Cells[1].CellStyle = style;
                row.CreateCell(2);
                row.Cells[2].SetCellValue("");
                row.Cells[2].CellStyle = style;
                row.CreateCell(3);//實際
                row.Cells[3].SetCellValue("實際數");
                row.Cells[3].CellStyle = style;
                row.CreateCell(4);//1月
                if (null != item.JAN && item.JAN.ToString().Trim() != "")
                {
                    row.Cells[4].SetCellValue(double.Parse(item.JAN.ToString()));
                }
                else
                {
                    row.Cells[4].SetCellValue("");
                }
                row.Cells[4].CellStyle = styleNumber;
                row.CreateCell(5);//2月
                if (null != item.FEB && item.FEB.ToString().Trim() != "")
                {
                    row.Cells[5].SetCellValue(double.Parse(item.FEB.ToString()));
                }
                else
                {
                    row.Cells[5].SetCellValue("");
                }
                row.Cells[5].CellStyle = styleNumber;
                row.CreateCell(6);//3月
                if (null != item.MAR && item.MAR.ToString().Trim() != "")
                {
                    row.Cells[6].SetCellValue(double.Parse(item.MAR.ToString()));
                }
                else
                {
                    row.Cells[6].SetCellValue("");
                }
                row.Cells[6].CellStyle = styleNumber;
                row.CreateCell(7);//4月
                if (null != item.APR && item.APR.ToString().Trim() != "")
                {
                    row.Cells[7].SetCellValue(double.Parse(item.APR.ToString()));
                }
                else
                {
                    row.Cells[7].SetCellValue("");
                }
                row.Cells[7].CellStyle = styleNumber;
                row.CreateCell(8);//5月
                if (null != item.MAY && item.MAY.ToString().Trim() != "")
                {
                    row.Cells[8].SetCellValue(double.Parse(item.MAY.ToString()));
                }
                else
                {
                    row.Cells[8].SetCellValue("");
                }
                row.Cells[8].CellStyle = styleNumber;
                row.CreateCell(9);//6月
                if (null != item.JUN && item.JUN.ToString().Trim() != "")
                {
                    row.Cells[9].SetCellValue(double.Parse(item.JUN.ToString()));
                }
                else
                {
                    row.Cells[9].SetCellValue("");
                }
                row.Cells[9].CellStyle = styleNumber;
                row.CreateCell(10);//7月
                if (null != item.JUL && item.JUL.ToString().Trim() != "")
                {
                    row.Cells[10].SetCellValue(double.Parse(item.JUL.ToString()));
                }
                else
                {
                    row.Cells[10].SetCellValue("");
                }
                row.Cells[10].CellStyle = styleNumber;
                row.CreateCell(11);//8月
                if (null != item.AUG && item.AUG.ToString().Trim() != "")
                {
                    row.Cells[11].SetCellValue(double.Parse(item.AUG.ToString()));
                }
                else
                {
                    row.Cells[11].SetCellValue("");
                }
                row.Cells[11].CellStyle = styleNumber;
                row.CreateCell(12);//9月
                if (null != item.SEP && item.SEP.ToString().Trim() != "")
                {
                    row.Cells[12].SetCellValue(double.Parse(item.SEP.ToString()));
                }
                else
                {
                    row.Cells[12].SetCellValue("");
                }
                row.Cells[12].CellStyle = styleNumber;
                row.CreateCell(13);//10月
                if (null != item.OCT && item.OCT.ToString().Trim() != "")
                {
                    row.Cells[13].SetCellValue(double.Parse(item.OCT.ToString()));
                }
                else
                {
                    row.Cells[13].SetCellValue("");
                }
                row.Cells[13].CellStyle = styleNumber;
                row.CreateCell(14);//11月
                if (null != item.NOV && item.NOV.ToString().Trim() != "")
                {
                    row.Cells[14].SetCellValue(double.Parse(item.NOV.ToString()));
                }
                else
                {
                    row.Cells[14].SetCellValue("");
                }
                row.Cells[14].CellStyle = styleNumber;
                row.CreateCell(15);//12月
                if (null != item.DEC && item.DEC.ToString().Trim() != "")
                {
                    row.Cells[15].SetCellValue(double.Parse(item.DEC.ToString()));
                }
                else
                {
                    row.Cells[15].SetCellValue("");
                }
                row.Cells[15].CellStyle = styleNumber;
                row.CreateCell(16);//合計
                if (null != item.HTOTAL && item.HTOTAL.ToString().Trim() != "")
                {
                    row.Cells[16].SetCellValue(double.Parse(item.HTOTAL.ToString()));
                }
                else
                {
                    row.Cells[16].SetCellValue("");
                }
                row.Cells[16].CellStyle = styleNumber;
                logger.Debug("get Site ExpenseByMonth cell style rowid=" + idxRow);
                idxRow++;
            }
            //4.另存新檔至專案所屬目錄 (增加Temp for zip 打包使用
            string fileLocation = null;
            if (isCum == true && targetMonth.ToString().Length > 1)
            {
                fileLocation = outputPath + "\\" + projectid + "\\" + projectid + "_" + targetYear + targetMonth + "_工地費用預算執行彙整表.xlsx";
            }
            else if (isCum == true)
            {
                fileLocation = outputPath + "\\" + projectid + "\\" + projectid + "_" + targetYear + "0" + targetMonth + "_工地費用預算執行彙整表.xlsx";
            }
            else
            {
                fileLocation = outputPath + "\\" + projectid + "\\" + projectid + "_" + targetYear + "_工地費用預算執行彙整表.xlsx";
            }
            var file = new FileStream(fileLocation, FileMode.Create);
            logger.Info("new file name =" + file.Name + ",path=" + file.Position);
            hssfworkbook.Write(file);
            file.Close();
            return fileLocation;
        }
        public SiteExpSummaryToExcel()
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

    #region 業主計價項目下載表格格式處理區段
    public class VAItem4OwnerToExcel
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        string wageFile = ContextService.strUploadPath + "\\wage_form.xlsx";
        string outputPath = ContextService.strUploadPath;

        IWorkbook hssfworkbook;
        ISheet sheet = null;
        string fileformat = "xlsx";
        //存放業主計價項目
        public TND_PROJECT project = null;
        //public List<PLAN_VALUATION_4OWNER> VAItems = null;
        public string errorMessage = null;
        string projId = null;
        //建立業主計價項目下載表格
        public void exportExcel(TND_PROJECT project, List<PROJECT_ITEM_WITH_WAGE> projectItems)
        {
            //1.讀取業主計價項目表格檔案
            InitializeWorkbook(wageFile);
            sheet = (XSSFSheet)hssfworkbook.GetSheet("工率");

            //2.填入表頭資料
            logger.Debug("Table Head_1=" + sheet.GetRow(1).Cells[0].ToString());
            sheet.GetRow(1).Cells[1].SetCellValue(project.PROJECT_ID);//專案編號
            logger.Debug("Table Head_2=" + sheet.GetRow(2).Cells[0].ToString());
            sheet.GetRow(2).Cells[1].SetCellValue(project.PROJECT_NAME);//專案名稱

            //3.填入表單明細
            int idxRow = 4;
            foreach (TND_PROJECT_ITEM item in projectItems)
            {
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                //PK(PROJECT_ITEM_ID) 項次 名稱 單位 數量 單價 備註 九宮格 次九宮格 主系統 次系統
                row.CreateCell(0).SetCellValue(item.PROJECT_ITEM_ID);//PK(PROJECT_ITEM_ID)
                row.CreateCell(1).SetCellValue(item.ITEM_ID);//項次
                logger.Debug("Wage :ITEM DESC=" + item.ITEM_DESC);
                row.CreateCell(2).SetCellValue(item.ITEM_DESC);//項目說明
                row.CreateCell(3).SetCellValue(item.ITEM_UNIT);// 單位
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.CreateCell(4).SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString())); //數量
                }
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "")
                {
                    row.CreateCell(5).SetCellValue(double.Parse(item.ITEM_UNIT_PRICE.ToString())); //單價
                }
                row.CreateCell(6).SetCellValue(item.ITEM_REMARK);// 備註
                row.CreateCell(7).SetCellValue(item.TYPE_CODE_1);// 九宮格
                row.CreateCell(8).SetCellValue(item.TYPE_CODE_2);// 次九宮格
                row.CreateCell(9).SetCellValue(item.SYSTEM_MAIN);// 主系統
                row.CreateCell(10).SetCellValue(item.SYSTEM_SUB);// 次系統
                //建立空白欄位
                for (int iTmp = 11; iTmp < 20; iTmp++)
                {
                    row.CreateCell(iTmp);
                }
                idxRow++;
            }
            //4.令存新檔至專案所屬目錄
            var file = new FileStream(outputPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "_工率.xlsx", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }
        public VAItem4OwnerToExcel()
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
        #region 業主計價項目資料轉換 
        /**
         * 取得業主計價項目Sheet 資料
         * */
        //public List<PLAN_VALUATION_4OWNER> ConvertDataForVAOfOwner(string projectId)
        //{
        //    projId = projectId;
        //    //1.依據檔案附檔名使用不同物件讀取Excel 檔案，並開啟整理後標單Sheet
        //    if (fileformat == "xls")
        //    {
        //        logger.Debug("office 2003:" + fileformat + " for projectID=" + projId + ":業主計價項目");
        //        sheet = (HSSFSheet)hssfworkbook.GetSheet("業主計價項目");
        //    }
        //    else
        //    {
        //        logger.Debug("office 2007:" + fileformat + " for projectID=" + projId + ":業主計價項目");
        //        sheet = (XSSFSheet)hssfworkbook.GetSheet("業主計價項目");
        //    }
        //    if (null == sheet)
        //    {
        //        logger.Error("檔案內沒有業主計價項目(Sheet)! filename=" + fileformat);
        //        throw new Exception("檔案內沒有[業主計價項目]資料");
        //    }
        //    return ConverData2VAOfOwner();
        //}
        ///**
        // * 轉換業主計價項目資料檔:計價項目
        // * */
        //protected List<PLAN_VALUATION_4OWNER> ConverData2VAOfOwner()
        //{
        //    IRow row = null;
        //    List<PLAN_VALUATION_4OWNER> lstItem = new List<PLAN_VALUATION_4OWNER>();
        //    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
        //    //2.逐行讀取資料
        //    int iRowIndex = 0; //0 表 Row 1

        //    //2.1  忽略不要的行數..(表頭)
        //    while (iRowIndex < (4))
        //    {
        //        rows.MoveNext();
        //        iRowIndex++;
        //        //row = (IRow)rows.Current;
        //        //logger.Debug("skip data Excel Value:" + row.Cells[0].ToString() + "," + row.Cells[1] + "," + row.Cells[2]);
        //    }
        //    //循序處理每一筆資料之欄位!!
        //    iRowIndex++;
        //    while (rows.MoveNext())
        //    {
        //        row = (IRow)rows.Current;
        //        int i = 0;
        //        string slog = "";
        //        for (i = 0; i < row.Cells.Count; i++)
        //        {
        //            slog = slog + "," + row.Cells[i];

        //        }
        //        logger.Debug("Excel Value:" + slog);
        //        //將各Row 資料寫入物件內
        //        //0.項次	1.工程項目 2.分項計價比率 3.說明(remark)
        //        if (row.Cells[0].ToString().ToUpper() != "END")
        //        {
        //            lstItem.Add(convertRow2VAOfOwner(row, iRowIndex));
        //        }
        //        else
        //        {
        //            logErrorMessage("Step1 ;取得業主計價項目資料:" + VAItems.Count + "筆");
        //            logger.Info("Finish convert Job : count=" + VAItems.Count);
        //            return lstItem;
        //        }
        //        iRowIndex++;
        //    }
        //    logger.Info("PLAN_VALUATION_4OWNER Count:" + iRowIndex);
        //    return lstItem;
        //}
        ///**
        // * 將Excel Row 轉換成為對應的資料物件
        // * */
        //private PLAN_VALUATION_4OWNER convertRow2VAOfOwner(IRow row, int excelrow)
        //{
        //    PLAN_VALUATION_4OWNER item = new PLAN_VALUATION_4OWNER();
        //    item.PROJECT_ID = projId;
        //    if (row.Cells[0].ToString().Trim() != "")//0.項次
        //    {
        //        item.ITEN_NO = row.Cells[0].ToString();
        //    }
        //    if (row.Cells[1].ToString().Trim() != "")//1.工程項目
        //    {
        //        item.ITEM_DESC = row.Cells[1].ToString();
        //    }
        //    if (null != row.Cells[2].ToString().Trim() || row.Cells[2].ToString().Trim() != "")//2.分項計價比率
        //    {
        //        try
        //        {
        //            decimal dRatio = decimal.Parse(row.Cells[2].ToString());
        //            logger.Info("excelrow=" + excelrow + ",value=" + row.Cells[2].ToString());
        //            item.ITEM_VALUATION_RATIO = dRatio;
        //        }
        //        catch (Exception e)
        //        {
        //            logger.Error("data format Error on ExcelRow=" + excelrow + ",Cells[12].value=" + row.Cells[2].ToString());
        //            logger.Error(e);
        //        }
        //    }
        //    if (row.Cells[3].ToString().Trim() != "")//3.說明
        //    {
        //        item.REMARK = row.Cells[3].ToString();
        //    }
        //    item.CREATE_DATE = System.DateTime.Now;
        //    logger.Info("PLAN_VALUATION_4OWNER=" + item.ToString());
        //    return item;
        //}

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
}
