﻿using log4net;
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
        public void exportExcel4po(PLAN_SUP_INQUIRY form, List<PLAN_SUP_INQUIRY_ITEM> formItems)
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
            //4.令存新檔至專案所屬目錄
            var file = new FileStream(outputPath + "\\" + form.PROJECT_ID + "\\" + ContextService.quotesFolder + "\\" + form.INQUIRY_FORM_ID + ".xlsx", FileMode.Create);
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
        public void convertInquiry2Plan(string fileExcel, string projectid)
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
            form.CONTACT_EMAIL = sheet.GetRow(6).Cells[1].ToString();
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
                        logger.Info("Project ITEM ID=" + row.Cells[row.Cells.Count - 1].ToString());
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