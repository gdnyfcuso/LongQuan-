using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Text.RegularExpressions;

namespace DBEN.DBI
{
    /// <summary>
    /// 从excel中将数据导出到datatable
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>读取excel
        /// 默认第一行为标头
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public static DataTable ImportExceltoDt(string strFileName)
        {
            var dt = new DataTable();
            IWorkbook wb;
            using (var file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                wb = WorkbookFactory.Create(file);
            }
            ISheet sheet = wb.GetSheetAt(0);
            dt = ImportDt(sheet, 1, true);
            return dt;
        }

        /// <summary>
        /// 读取Excel到table
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="headerRowIndex">列头所在行号，-1表示没有列头</param>
        /// <returns></returns>
        public static DataTable ImportExceltoDt(string filePath, int headerRowIndex)
        {
            var dt = new DataTable("ExcelTable");
            IWorkbook wb;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                wb = WorkbookFactory.Create(file);
            }
            ISheet sheet = wb.GetSheetAt(0);
            dt = ImportDt(sheet, headerRowIndex, true);
            return dt;
        }

        public static DataTable ImportExceltoDt(string filePath, int headerRowIndex, string type, int notNullCount = 0)
        {
            var dt = new DataTable("ExcelTable");
            IWorkbook wb;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                wb = WorkbookFactory.Create(file);
            }
            ISheet sheet = wb.GetSheetAt(0);
            dt = ImportDt_e(sheet, headerRowIndex, true, type, notNullCount);
            return dt;
        }

        /// <summary>
        /// 读取excel
        /// </summary>xdz 
        /// <param name="strFileName">excel文件路径</param>
        /// <param name="sheetName">需要导出的sheet</param>
        /// <param name="headerRowIndex">列头所在行号，-1表示没有列头</param>
        /// <param name="needHeader"></param>
        /// <returns></returns>
        public static DataTable ImportExceltoDt(string strFileName, string sheetName, int headerRowIndex, bool needHeader)
        {
            IWorkbook wb;
            using (var file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                wb = WorkbookFactory.Create(file);
            }
            ISheet sheet = wb.GetSheet(sheetName);
            var table = new DataTable();
            table = ImportDt(sheet, headerRowIndex, needHeader);
            //ExcelFileStream.Close();
            sheet = null;
            return table;
        }

        /// <summary>
        /// 将制定sheet中的数据导出到datatable中
        /// </summary>
        /// <param name="sheet">需要导出的sheet</param>
        /// <param name="headerRowIndex">列头所在行号，-1表示没有列头</param>
        /// <param name="needHeader"></param>
        /// <param name="tableNameRowIndex"></param>
        /// <returns></returns>
        static DataTable ImportDt(ISheet sheet, int headerRowIndex, bool needHeader, int tableNameRowIndex = -1)
        {
            var table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                if (headerRowIndex < 0 || !needHeader)
                {
                    headerRow = sheet.GetRow(0);
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        var column = new DataColumn(Convert.ToString(i));
                        table.Columns.Add(column);
                    }
                }
                else
                {
                    headerRow = sheet.GetRow(headerRowIndex);
                    cellCount = headerRow.LastCellNum;
                    if (tableNameRowIndex > -1)
                    {
                        var head = sheet.GetRow(tableNameRowIndex);
                        table.TableName = head.GetCell(0).StringCellValue;
                    }

                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        if (headerRow.GetCell(i) == null)
                        {
                            if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                            {
                                var column = new DataColumn(Convert.ToString("重复列名" + i));
                                table.Columns.Add(column);
                            }
                            else
                            {
                                var column = new DataColumn(Convert.ToString(i));
                                table.Columns.Add(column);
                            }
                        }
                        else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                        {
                            var column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            var column = new DataColumn(headerRow.GetCell(i).ToString());
                            table.Columns.Add(column);
                        }
                    }
                }
                int rowCount = sheet.LastRowNum;
                for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }

                        DataRow dataRow = table.NewRow();

                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.String:
                                            string str = row.GetCell(j).StringCellValue;
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = null;
                                            }
                                            break;
                                        case CellType.Numeric:
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.Error:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.Formula:
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    string strFORMULA = row.GetCell(j).StringCellValue;
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                var str = exception.Message;
                                throw;
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {
                        var str = exception.Message;
                        throw;
                    }
                }
            }
            catch (Exception exception)
            {
                var str = exception.Message;
                throw;
            }
            return table;
        }

        /// <summary>
        /// DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">保存位置</param>
        public static void DataTableToExcel(DataTable dtSource, string strHeaderText, string strFileName)
        {
            string[] temp = strFileName.Split('.');

            if (temp[temp.Length - 1] == "xls" && dtSource.Columns.Count < 256 && dtSource.Rows.Count < 65536)
            {
                using (MemoryStream ms = ExportDt(dtSource, strHeaderText))
                {
                    using (var fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                }
            }
            else
            {
                if (temp[temp.Length - 1] == "xls")
                    strFileName = strFileName + "x";

                using (var fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    ExportDti(dtSource, strHeaderText, fs);
                }

            }
        }

        /// <summary>
        /// DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        public static MemoryStream ExportDt(DataTable dtSource, string strHeaderText)
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet() as HSSFSheet;

            var cellStyle = workbook.CreateBasicCellStyle();

            HSSFCellStyle dateStyle = workbook.CreateBasicCellStyle();
            var format = workbook.CreateDataFormat() as HSSFDataFormat;
            // dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd HH:mm:ss");
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");
            //取得列宽
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName).Length + 6;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;

            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet() as HSSFSheet;
                    }

                    #region 表头及样式

                    if (!string.IsNullOrEmpty(strHeaderText) && strHeaderText != "下载开票模板_Evan")
                    {
                        var headerRow = sheet.CreateRow(0) as HSSFRow;
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        var headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        var font = workbook.CreateFont() as HSSFFont;
                        font.FontHeightInPoints = 16;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);

                        headerRow.GetCell(0).CellStyle = headStyle;

                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                        //headerRow.Dispose();

                        if (dtSource.Columns[dtSource.Columns.Count - 1].ColumnName == "编号")
                        {
                            sheet.SetColumnHidden(dtSource.Columns.Count - 1, true);
                        }
                        //  sheet.SetColumnHidden();

                    }

                    #endregion

                    #region 列头及样式

                    {
                        var rownum = 0;
                        if (strHeaderText == "下载开票模板_Evan")
                        {
                            rownum = 0;
                        }
                        else
                        {
                            rownum = string.IsNullOrEmpty(strHeaderText) ? 0 : 1;
                        }
                        var headerRow = sheet.CreateRow(rownum) as HSSFRow;

                        headerRow.HeightInPoints = 20;

                        var headStyle = workbook.CreateCellStyle() as HSSFCellStyle;
                        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        var font = workbook.CreateFont() as HSSFFont;
                        font.FontHeightInPoints = 11;
                        font.Boldweight = 600;
                        headStyle.SetFont(font);

                        headStyle.BorderLeft = BorderStyle.Thin;
                        headStyle.BorderRight = BorderStyle.Thin;
                        headStyle.BorderTop = BorderStyle.Thin;
                        headStyle.BorderBottom = BorderStyle.Thin;

                        foreach (DataColumn column in dtSource.Columns)
                        {
                            if (strHeaderText == "下载开票模板_Evan")
                            {
                                //headStyle.IsLocked = true;
                                headerRow.CreateCell(column.Ordinal).CellStyle = headStyle;
                                sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            }
                            else
                            {
                                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                                headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                                //设置列宽
                                //sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                            }
                        }
                        //sheet.CreateFreezePane(0, 1, 0, 1);
                        //sheet.CreateFreezePane(1, 0, 1, 0);
                        //headerRow.Dispose();
                    }

                    #endregion

                    if (strHeaderText == "下载开票模板_Evan")
                    {
                        rowIndex = 1;
                    }
                    else
                    {
                        rowIndex = string.IsNullOrEmpty(strHeaderText) ? 1 : 2;
                    }
                }

                #endregion

                #region 填充内容

                var dataRow = sheet.CreateRow(rowIndex) as HSSFRow;

                foreach (DataColumn column in dtSource.Columns)
                {
                    var newCell = dataRow.CreateCell(column.Ordinal) as HSSFCell;

                    newCell.CellStyle = cellStyle;

                    string drValue = row[column].ToString();

                    #region 写单元格的值

                    switch (column.DataType.ToString())
                    {
                        case "System.String": //字符串类型
                            double result;
                            if (IsNumeric(drValue, out result))
                            {
                                double.TryParse(drValue, out result);
                                newCell.SetCellValue(result);
                                break;
                            }
                            else
                            {
                                newCell.SetCellValue(drValue);
                                break;
                            }

                        case "System.DateTime": //日期类型
                            DateTime dateV;
                            if (DateTime.TryParse(drValue, out dateV))
                            {
                                newCell.SetCellValue(dateV);
                                newCell.CellStyle = dateStyle; //格式化显示
                            }
                            break;
                        case "System.Boolean": //布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16": //整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal": //浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull": //空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue(drValue);
                            break;
                    }

                    #endregion

                    if (strHeaderText == "下载开票模板_Evan")
                    {
                        if (column.ColumnName == "Booking Branch" || column.ColumnName == "Customer ID" || column.ColumnName == "EBBS relationship ID" || column.ColumnName == "Customer ID/counterparty ID" || column.ColumnName == "EBBS Master ID" || column.ColumnName == "分行代码" || column.ColumnName == "客户编码")
                        {
                            newCell.SetCellValue(drValue);
                        }
                    }
                }

                #endregion

                rowIndex++;
            }
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;

                //sheet.Dispose();
                //workbook.Dispose();

                return ms;
            }
        }

        private static HSSFCellStyle CreateBasicCellStyle(this HSSFWorkbook workbook)
        {
            var cellStyle = workbook.CreateCellStyle() as HSSFCellStyle;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            return cellStyle;
        }

        /// <summary>
        /// DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="fs">The fs.</param>
        public static void ExportDti(DataTable dtSource, string strHeaderText, FileStream fs)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet() as XSSFSheet;

            var dateStyle = workbook.CreateCellStyle() as XSSFCellStyle;
            var format = workbook.CreateDataFormat() as XSSFDataFormat;
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            //取得列宽
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;

            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 0)
                {

                    #region 列头及样式

                    {
                        var headerRow = sheet.CreateRow(0) as XSSFRow;


                        var headStyle = workbook.CreateCellStyle() as XSSFCellStyle;
                        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        var font = workbook.CreateFont() as XSSFFont;
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);


                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);

                        }
                        //headerRow.Dispose();
                    }

                    #endregion

                    rowIndex = 1;
                }

                #endregion

                #region 填充内容

                var dataRow = sheet.CreateRow(rowIndex) as XSSFRow;
                foreach (DataColumn column in dtSource.Columns)
                {
                    var newCell = dataRow.CreateCell(column.Ordinal) as XSSFCell;

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String": //字符串类型
                            double result;
                            if (IsNumeric(drValue, out result))
                            {

                                double.TryParse(drValue, out result);
                                newCell.SetCellValue(result);
                                break;
                            }
                            else
                            {
                                newCell.SetCellValue(drValue);
                                break;
                            }

                        case "System.DateTime": //日期类型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle; //格式化显示
                            break;
                        case "System.Boolean": //布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16": //整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal": //浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull": //空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }

                }

                #endregion

                rowIndex++;
            }
            workbook.Write(fs);
            fs.Close();
        }

        private static bool IsNumeric(string message, out double result)
        {
            var rex = new Regex(@"^[-]?\d+[.]?\d*$");
            result = -1;
            if (rex.IsMatch(message))
            {
                result = double.Parse(message);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 如何判断是不是汉字
        /// </summary>
        /// <param name="cString"></param>
        /// <returns></returns>
        public static bool IsChinaString(string cString)
        {
            bool BoolValue = false;
            for (int i = 0; i < cString.Length; i++)
            {
                if (Convert.ToInt32(Convert.ToChar(cString.Substring(i, 1))) < Convert.ToInt32(Convert.ToChar(128)))
                {
                    BoolValue = false;
                }
                else
                {
                    return BoolValue = true;
                }
            }
            return BoolValue;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headerRowIndex"></param>
        /// <param name="needHeader"></param>
        /// <param name="type"></param>
        /// <param name="notNullCount">至少有几列不为空</param>
        /// <param name="tableNameRowIndex"></param>
        /// <returns></returns>
        static DataTable ImportDt_e(ISheet sheet, int headerRowIndex, bool needHeader, string type, int notNullCount = 0, int tableNameRowIndex = -1)
        {
            var table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                if (headerRowIndex < 0 || !needHeader)
                {
                    headerRow = sheet.GetRow(0);
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        var column = new DataColumn(Convert.ToString(i));
                        table.Columns.Add(column);
                    }
                }
                else
                {
                    headerRow = sheet.GetRow(headerRowIndex);
                    cellCount = headerRow.LastCellNum;
                    if (tableNameRowIndex > -1)
                    {
                        var head = sheet.GetRow(tableNameRowIndex);
                        table.TableName = head.GetCell(0).StringCellValue;
                    }

                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        if (headerRow.GetCell(i) == null)
                        {
                            if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                            {
                                var column = new DataColumn(Convert.ToString("重复列名" + i));
                                table.Columns.Add(column);
                            }
                            else
                            {
                                var column = new DataColumn(Convert.ToString(i).Replace(" ", ""));
                                if (Convert.ToString(i) != "")
                                {
                                    table.Columns.Add(column);
                                }
                            }
                        }
                        else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                        {
                            var column = new DataColumn(Convert.ToString("重复列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        {
                            var column = new DataColumn(headerRow.GetCell(i).ToString().Replace(" ", ""));
                            if (Convert.ToString(i) != "")
                            {
                                table.Columns.Add(column);
                            }
                        }
                    }
                }
                int rowCount = sheet.LastRowNum;
                for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }

                        DataRow dataRow = table.NewRow();
                        var nullColums = 0;
                        for (int j = headerRow.FirstCellNum; j < cellCount; j++)
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    if (type == "0" && row.GetCell(j).CellType == CellType.String && j == 21)
                                    {
                                        double dub = 0;
                                        if (double.TryParse(row.GetCell(j).StringCellValue.Replace("%", ""), out dub))
                                        {
                                            dataRow[j] = Convert.ToDouble(dub / 100);
                                        }
                                        else
                                        {
                                            dataRow[j] = row.GetCell(j).StringCellValue;
                                        }
                                    }
                                    else
                                    {
                                        switch (row.GetCell(j).CellType)
                                        {
                                            case CellType.String:
                                                string str = row.GetCell(j).StringCellValue;
                                                if (str != null && str.Length > 0)
                                                {
                                                    dataRow[j] = str.ToString();
                                                }
                                                else
                                                {
                                                    dataRow[j] = null;
                                                }
                                                break;
                                            case CellType.Numeric:
                                                if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                                {
                                                    dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                                }
                                                else
                                                {
                                                    dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                                }
                                                break;
                                            case CellType.Boolean:
                                                dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                break;
                                            case CellType.Error:
                                                dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                break;
                                            case CellType.Formula:
                                                switch (row.GetCell(j).CachedFormulaResultType)
                                                {
                                                    case CellType.String:
                                                        string strFORMULA = row.GetCell(j).StringCellValue;
                                                        if (strFORMULA != null && strFORMULA.Length > 0)
                                                        {
                                                            dataRow[j] = strFORMULA.ToString();
                                                        }
                                                        else
                                                        {
                                                            dataRow[j] = null;
                                                        }
                                                        break;
                                                    case CellType.Numeric:
                                                        dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                        break;
                                                    case CellType.Boolean:
                                                        dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                        break;
                                                    case CellType.Error:
                                                        dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                        break;
                                                    default:
                                                        dataRow[j] = "";
                                                        break;
                                                }
                                                break;
                                            default:
                                                dataRow[j] = "";
                                                break;
                                        }
                                    }
                                }
                                if (dataRow[j].ToString() == "")
                                {
                                    nullColums++;
                                }
                            }
                            catch (Exception exception)
                            {
                                var str = exception.Message;
                                throw;
                            }
                        }
                        if (nullColums < cellCount - notNullCount)//至少有几列是有值的
                        {
                            table.Rows.Add(dataRow);
                        }
                        else
                        {
                            //i = sheet.LastRowNum;
                            continue;//断行的下一行，可能还有会数据，需要继续判断
                        }
                    }
                    catch (Exception exception)
                    {
                        var str = exception.Message;
                        throw;
                    }
                }
            }
            catch (Exception exception)
            {
                var str = exception.Message;
                throw;
            }
            return table;
        }
    }
}
