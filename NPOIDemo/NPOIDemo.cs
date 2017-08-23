using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIDemo
{
    public class NPOIDemo
    {
        #region 创建Excel颜色列表
        public static void CreateExcelColorFile(string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            
            // 创建Workbook。
            NPOI.SS.UserModel.IWorkbook workbook = null;// new NPOI.XSSF.UserModel.XSSFWorkbook();
            if (fileInfo.Extension == ".xlsx")
            {
                workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            }
            else if (fileInfo.Extension == ".xls")
            {
                workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            }
            else
            {
                return;
            }
            // 创建Sheet
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("ExcelColor");
            // 设置列宽。
            sheet.SetColumnWidth(0, 15 * 256);
            sheet.SetColumnWidth(1, 15 * 256);
            sheet.SetColumnWidth(2, 20 * 256);
            sheet.SetColumnWidth(3, 20 * 256);
            sheet.SetColumnWidth(4, 15 * 256);
            // 创建标题。
            int rowIndex = 0;
            NPOI.SS.UserModel.ICell cell = sheet.CreateRow(rowIndex).CreateCell(0);
            cell.SetCellValue("Excel颜色");
            cell.CellStyle = NPOIExtension.GetCellStyle(workbook, ExcelColor.None.IndexNPOI, ExcelColor.Black.IndexNPOI, new System.Drawing.Font("Arial", 16, System.Drawing.FontStyle.Bold));
            // 空一行
            rowIndex++;
            // 创建标题行
            rowIndex++;
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(rowIndex);
            row.Height = 30 * 20;
            NPOI.SS.UserModel.ICellStyle columnHeadCellStyle = NPOIExtension.GetCellStyle(workbook, ExcelColor.None.IndexNPOI, ExcelColor.Black.IndexNPOI, new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.VerticalAlignment.Center);
            cell = row.CreateCell(0);
            cell.SetCellValue("索引");
            cell.CellStyle = columnHeadCellStyle;

            cell = row.CreateCell(1);
            cell.SetCellValue("索引（NPOI）");
            cell.CellStyle = columnHeadCellStyle;

            cell = row.CreateCell(2);
            cell.SetCellValue("名称");
            cell.CellStyle = columnHeadCellStyle;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 2, 3));

            cell = row.CreateCell(4);
            cell.SetCellValue("十六进制代码");
            cell.CellStyle = columnHeadCellStyle;

            ExcelColor.KnownColors.ForEach(x =>
            {
                rowIndex++;
                NPOI.SS.UserModel.IRow contentRow = sheet.CreateRow(rowIndex);
                contentRow.Height = 30 * 20;
                NPOI.SS.UserModel.ICellStyle contentCellStyle = NPOIExtension.GetCellStyle(workbook, x.IndexNPOI, (short)(x.IsDarkColor ? 9 : 8), new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.VerticalAlignment.Center);
                NPOI.SS.UserModel.ICell contentCell = contentRow.CreateCell(0);
                contentCell.SetCellValue(x.Index);
                contentCell.CellStyle = contentCellStyle;

                contentCell = contentRow.CreateCell(1);
                contentCell.SetCellValue(x.IndexNPOI + "（NPOI）");
                contentCell.CellStyle = contentCellStyle;

                contentCell = contentRow.CreateCell(2);
                contentCell.SetCellValue(x.Name);
                contentCell.CellStyle = NPOIExtension.GetCellStyle(workbook, x.IndexNPOI, (short)(x.IsDarkColor ? 9 : 8), new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Right, NPOI.SS.UserModel.VerticalAlignment.Center);

                contentCell = contentRow.CreateCell(3);
                contentCell.SetCellValue(x.Description);
                contentCell.CellStyle = NPOIExtension.GetCellStyle(workbook, x.IndexNPOI, (short)(x.IsDarkColor ? 9 : 8), new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Center);

                contentCell = contentRow.CreateCell(4);
                contentCell.SetCellValue(x.HexString);
                contentCell.CellStyle = contentCellStyle;

            });
            // 保存到文件。
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite);
            workbook.Write(fileStream);
            fileStream.Close();
            workbook.Close();

        }
        #endregion

        #region 读取Excel
        public static System.Data.DataTable GetDataTabelFromExcelFile(string fileName, int startRowIndex)
        {
            System.Data.DataTable dataSource = new System.Data.DataTable();
            dataSource.Columns.Add("Index", typeof(string));
            dataSource.Columns.Add("IndexNPOI", typeof(string));
            dataSource.Columns.Add("Name", typeof(string));
            dataSource.Columns.Add("Description", typeof(string));
            dataSource.Columns.Add("HexString", typeof(string));

            // 获取Workbook。
            NPOI.SS.UserModel.IWorkbook workbook = NPOIExtension.GetWorkbookFromExcelFile(fileName);
            if (workbook == null)
            { return dataSource; }

            // 获取Sheet
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            for (int i = startRowIndex; i < sheet.LastRowNum; i++)
            {
                NPOI.SS.UserModel.IRow row = sheet.GetRow(i);
                System.Data.DataRow dataSourceRow = dataSource.NewRow();
                dataSource.Rows.Add(dataSourceRow);
                for (int j = 0; j < dataSource.Columns.Count; j++)
                {
                    dataSourceRow[j] = row.GetCell(j).ToString();
                }
            }
            return dataSource;
        }
        #endregion
    }
}
