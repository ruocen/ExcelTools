using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace ExcelTools
{
    public class ExcelRows
    {
        public string Code { get; set; }
        public List<string> ExcelCells { get; set; }
    }
    public class ExcelHelper
    {
        public static void Import(string fileDialogFileName, string fileDialogSafeFileNames, out string FileName,out string fileExt, out List<string> headers, out List<ExcelRows> rowlist, out Dictionary<string, int> link, string type)
        {
            link = new Dictionary<string, int>();
            headers = new List<string>();
            rowlist = new List<ExcelRows>();
            IWorkbook workbook;
            fileExt = Path.GetExtension(fileDialogFileName).ToLower();
            using (FileStream file = new FileStream(fileDialogFileName, FileMode.Open, FileAccess.Read))
            {
                if (fileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook(file);
                    FileName = fileDialogSafeFileNames.Substring(0, fileDialogSafeFileNames.Length - 5);
                }
                else if (fileExt == ".xls")
                {
                    workbook = new HSSFWorkbook(file);
                    FileName = fileDialogSafeFileNames.Substring(0, fileDialogSafeFileNames.Length - 4);
                }
                else
                {
                    workbook = null;
                    FileName = "";
                    link = null;
                    headers = null;
                    rowlist = null;
                }
                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet == null)
                {
                    FileName = "";
                    link = null;
                    headers = null;
                    rowlist = null;
                }
                //表头  
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    if (i == 0 && type == "right") continue;
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        headers.Add("Columns" + i.ToString());
                    }
                    else
                        headers.Add(obj.ToString());
                }

                //数据  
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    ExcelRows excelRows = new ExcelRows();
                    List<string> list = new List<string>();
                    for (int j = sheet.GetRow(i).FirstCellNum; j < sheet.GetRow(i).LastCellNum; j++)
                    {
                        var cellValue = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (j == 2 && type == "left")
                            excelRows.Code = cellValue.ToString().Trim().ToLower();
                        if (j == 0 && type == "right")
                        {
                            excelRows.Code = cellValue.ToString().Trim().ToLower();
                            link.Add(excelRows.Code, i - 1);
                        }
                        if (cellValue != null && cellValue.ToString() != string.Empty)
                            list.Add(cellValue.ToString().Trim().ToLower());
                    }
                    excelRows.ExcelCells = list;
                    rowlist.Add(excelRows);
                }
            }
        }

       
        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}
