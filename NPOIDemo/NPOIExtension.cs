using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIDemo
{
    public static class NPOIExtension
    {
        #region 获取单元格样式
        public static NPOI.SS.UserModel.ICellStyle GetCellStyle(this NPOI.SS.UserModel.IWorkbook workbook, short backColorIndex, short fontColorIndex, System.Drawing.Font font, NPOI.SS.UserModel.HorizontalAlignment horizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.VerticalAlignment verticalAlignment = NPOI.SS.UserModel.VerticalAlignment.None, NPOI.SS.UserModel.BorderStyle borderLeft = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderTop = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderRight = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderBottom = NPOI.SS.UserModel.BorderStyle.None)
        {
            NPOI.SS.UserModel.ICellStyle cellStyle = workbook.CreateCellStyle();
            if (backColorIndex >= 8 && backColorIndex <= 63)
            {
                cellStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                cellStyle.FillForegroundColor = backColorIndex;
            }
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.VerticalAlignment = verticalAlignment;
            cellStyle.BorderLeft = borderLeft;
            cellStyle.BorderTop = borderTop;
            cellStyle.BorderRight = borderRight;
            cellStyle.BorderBottom = borderBottom;

            if (font != null)
            {
                NPOI.SS.UserModel.IFont cellFont = workbook.CreateFont();
                if (fontColorIndex >= 8 && fontColorIndex <= 63)
                { cellFont.Color = fontColorIndex; }
                cellFont.FontName = font.Name;
                cellFont.FontHeight = font.Size;
                cellFont.Boldweight = (short)(font.Bold ? NPOI.SS.UserModel.FontBoldWeight.Bold : NPOI.SS.UserModel.FontBoldWeight.Normal);
                cellStyle.SetFont(cellFont);
            }

            return cellStyle;
        }
        #endregion

        #region 将工作簿保存到文件
        public static void SaveToFile(NPOI.SS.UserModel.IWorkbook workbook, string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite);
            workbook.Write(fileStream);
            fileStream.Close();
            workbook.Close();
        }
        #endregion

        #region 从文件中获取工作簿
        public static NPOI.SS.UserModel.IWorkbook GetWorkbookFromExcelFile(string fileName)
        {
            NPOI.SS.UserModel.IWorkbook workbook = null;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            if (!fileInfo.Exists)
            { return workbook; }
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            if (fileInfo.Extension == ".xlsx")
            { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fileStream); }
            else if (fileInfo.Extension == ".xls")
            { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fileStream); }
            fileStream.Close();
            return workbook;
        }
        #endregion

        #region 从文件中获取数据
        public static System.Data.DataTable GetDataTabelFromExcelFile(string fileName, bool firstLineIsColumnHead)
        {
            System.Data.DataTable dataSource = new System.Data.DataTable();
            NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
            if (workbook == null)
            { return dataSource; }
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(workbook.ActiveSheetIndex);
            if (sheet == null)
            { return dataSource; }
            NPOI.SS.UserModel.IRow firstRow = sheet.GetRow(0);
            if (firstRow != null)
            {
                for (int columnIndex = 0; columnIndex < firstRow.LastCellNum; columnIndex++)
                {
                    if (firstLineIsColumnHead)
                    {
                        NPOI.SS.UserModel.ICell cell = firstRow.GetCell(columnIndex);
                        if (cell == null)
                        {
                            dataSource.Columns.Add();
                        }
                        else
                        {
                            dataSource.Columns.Add(cell.ToString());
                        }
                    }
                    else
                    {
                        dataSource.Columns.Add();
                    }
                }
            }
            int startRowIndex = firstLineIsColumnHead ? 1 : 0;
            for (int rowIndex = startRowIndex; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                System.Data.DataRow dataSourceRow = dataSource.NewRow();
                dataSource.Rows.Add(dataSourceRow);
                NPOI.SS.UserModel.IRow row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    for (int columnIndex = 0; columnIndex < dataSource.Columns.Count; columnIndex++)
                    {
                        NPOI.SS.UserModel.ICell cell = row.GetCell(columnIndex);
                        if (cell != null)
                        {
                            switch (cell.CellType)
                            {
                                case NPOI.SS.UserModel.CellType.Blank:
                                    dataSourceRow[columnIndex] = "";
                                    break;
                                case NPOI.SS.UserModel.CellType.Boolean:
                                    dataSourceRow[columnIndex] = cell.BooleanCellValue;
                                    break;
                                case NPOI.SS.UserModel.CellType.Error:
                                    dataSourceRow[columnIndex] = cell.ErrorCellValue;
                                    break;
                                case NPOI.SS.UserModel.CellType.Formula:
                                    dataSourceRow[columnIndex] = cell.NumericCellValue;
                                    break;
                                case NPOI.SS.UserModel.CellType.Numeric:
                                    if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(cell))
                                    { dataSourceRow[columnIndex] = cell.DateCellValue; }
                                    else
                                    { dataSourceRow[columnIndex] = cell.NumericCellValue; }
                                    break;
                                case NPOI.SS.UserModel.CellType.String:
                                    dataSourceRow[columnIndex] = cell.StringCellValue;
                                    break;
                                default:
                                    dataSourceRow[columnIndex] = cell.ToString();
                                    break;
                            }
                            
                        }
                        
                    }
                }
            }

            return dataSource;
        }
        #endregion

        #region 从文件中获取数据
        public static System.Data.DataTable GetDataTabelFromExcelFile(string fileName, int sheetNumber, int startColumnNumber, int startRowNumber, params string[] columnNames)
        {
            System.Data.DataTable dataSource = new System.Data.DataTable();
            if (columnNames.Length < 1)
            { return dataSource; }
            foreach (string columnName in columnNames)
            { dataSource.Columns.Add(columnName); }
            NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
            if (workbook == null)
            { return dataSource; }
            if (sheetNumber > workbook.NumberOfSheets)
            { return dataSource; }
            int sheetIndex = sheetNumber > 0 ? sheetNumber -1 : workbook.ActiveSheetIndex;
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(sheetIndex);
            int startRowIndex = startRowNumber > 0 ? startRowNumber - 1 : 0;
            for (int i = startRowIndex; i <= sheet.LastRowNum; i++)
            {
                System.Data.DataRow dataSourceRow = dataSource.NewRow();
                dataSource.Rows.Add(dataSourceRow);
                NPOI.SS.UserModel.IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    int startColumnIndex = startColumnNumber > 0 ? startColumnNumber - 1 : 0;
                    int columnIndex = 0;
                    Enumerable.Range(startColumnIndex, dataSource.Columns.Count).ToList().ForEach(cellIndex =>
                    {
                        NPOI.SS.UserModel.ICell cell = row.GetCell(cellIndex);
                        if (cell != null)
                        {
                            //switch (cell.CellType)
                            //{
                            //    case NPOI.SS.UserModel.CellType.Blank:
                            //        dataSourceRow[columnIndex] = "";
                            //        break;
                            //}
                            dataSourceRow[columnIndex] = cell.ToString();
                        }
                        columnIndex++;
                    });
                }
            }
            return dataSource;
        }
        #endregion

    }
}
