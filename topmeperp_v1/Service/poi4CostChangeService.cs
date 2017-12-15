using log4net;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using topmeperp.Models;

namespace topmeperp.Service
{
    public class poi4CostChangeService : ExcelBase
    {
        static ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public poi4CostChangeService()
        {
            //定義樣板檔案名稱
            templateFile = strUploadPath + "\\CostChange_Template_v1.xlsx";
            logger.Debug("Constroctor!!" + templateFile);
        }
        public void createExcel(TND_PROJECT project, PLAN_COSTCHANGE_FORM form, List<PLAN_COSTCHANGE_ITEM> lstItem)
        {
            InitializeWorkbook();
            SetOpSheet("異動單");
            //填寫專案資料
            IRow row = sheet.GetRow(1);
            row.Cells[1].SetCellValue(project.PROJECT_ID);
            row.Cells[2].SetCellValue(project.PROJECT_NAME);
            //填入異動單資料
            row = sheet.GetRow(2);
            row.Cells[1].SetCellValue(form.FORM_ID);
            row.Cells[3].SetCellValue(form.REMARK);
            //填入明細資料
            ConvertExcelToObject(lstItem, 4);
            //令存新檔至專案所屬目錄
            outputFile = strUploadPath + "\\" + project.PROJECT_ID + "\\" + project.PROJECT_ID + "-" + form.FORM_ID + "_CostChange.xlsx";
            logger.Debug("export excel file=" + outputFile);
            var file = new FileStream(outputFile, FileMode.Create);
            logger.Info("output file=" + file.Name);
            hssfworkbook.Write(file);
            file.Close();
        }
        //轉換物件
        public void ConvertExcelToObject(List<PLAN_COSTCHANGE_ITEM> lstItem, int startrow)
        {
            int idxRow = 5;

            foreach (PLAN_COSTCHANGE_ITEM item in lstItem)
            {
                logger.Debug("Row Id=" + idxRow + "," + item.ITEM_DESC);
                IRow row = sheet.CreateRow(idxRow);//.GetRow(idxRow);
                //編號 標單編號 項次 品項名稱 單位 單價 異動數量 備註說明 轉入標單
                row.CreateCell(0).SetCellValue(item.ITEM_UID);//PK(PROJECT_ITEM_ID)
                row.Cells[0].CellStyle = style;
                row.CreateCell(1).SetCellValue(item.PLAN_ITEM_ID);//標單編號
                row.Cells[1].CellStyle = style;
                row.CreateCell(2).SetCellValue(item.ITEM_ID);//項次
                row.Cells[2].CellStyle = style;
                row.CreateCell(3).SetCellValue(item.ITEM_DESC);// 品項名稱
                row.Cells[3].CellStyle = style;

                row.CreateCell(4).SetCellValue(item.ITEM_UNIT);//單位
                row.Cells[4].CellStyle = style;


                //異動數量
                if (null != item.ITEM_QUANTITY && item.ITEM_QUANTITY.ToString().Trim() != "")
                {
                    row.CreateCell(5).SetCellValue(double.Parse(item.ITEM_QUANTITY.ToString()));
                    row.Cells[5].CellStyle = style;
                }
                else
                {
                    row.CreateCell(5).SetCellValue("");
                }

                //單價 (還沒決定)
                ICell cel6 = row.CreateCell(6);
                if (null != item.ITEM_UNIT_PRICE && item.ITEM_UNIT_PRICE.ToString().Trim() != "")
                {
                    logger.Debug("UNIT PRICE=" + item.ITEM_UNIT_PRICE);
                    cel6.SetCellValue(double.Parse(item.ITEM_UNIT_PRICE.ToString()));
                    cel6.CellStyle = styleNumber;
                }
                else
                {
                    cel6.SetCellValue("");
                    cel6.CellStyle = styleNumber;
                }
                //複價
                ICell cel7 = row.CreateCell(7);
                if (null != item.ITEM_QUANTITY && null != item.ITEM_UNIT_PRICE)
                {
                    logger.Debug("Fomulor=" + "F" + (idxRow + 1) + "*G" + (idxRow + 1));
                    cel7.CellFormula = "F" + (idxRow + 1) + "*G" + (idxRow + 1);
                    cel7.CellStyle = styleNumber;
                }
                else
                {
                    cel7.SetCellValue("");
                    cel7.CellStyle = styleNumber;
                }
                //8 備註
                if (null != item.ITEM_REMARK && item.ITEM_REMARK.ToString().Trim() != "")
                {
                    row.CreateCell(8).SetCellValue(item.ITEM_REMARK);
                    row.Cells[8].CellStyle = style;
                }
                else
                {
                    row.CreateCell(8).SetCellValue("");
                }
                //9 追加/轉入標單
                if (null != item.TRANSFLAG && item.TRANSFLAG.ToString().Trim() != "")
                {
                    row.CreateCell(9).SetCellValue(item.TRANSFLAG);
                    row.Cells[9].CellStyle = style;
                }
                else
                {
                    row.CreateCell(9).SetCellValue("N");
                }

                idxRow++;
            }
            logger.Info("InitialQuotation finish!!");
        }
    }
}
