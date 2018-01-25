using Microsoft.Win32;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ETL.Utils
{
    public class Tools
    {
        #region excelを保存し、ファイル名をreturn
        public static string SaveExcelFileDialog()
        {
            var sfd = new SaveFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "2007excel files(*.xlsx)|*.xlsx|csv files(*.csv)|*.csv|All files(*.*)|*.*",
                FilterIndex = 1
            };

            if (sfd.ShowDialog() == false)
                return "";
            return sfd.FileName;
        }
        #endregion

        #region excelを開き、ファイル名をreturn
        public static string OpenExcelFileDialog()
        {
            var ofd = new OpenFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "2007excel files(*.xlsx)|*.xlsx|csv files(*.csv)|*.csv|All files(*.*)|*.*",
                FilterIndex = 1
            };

            if (ofd.ShowDialog() == false)
                return "";
            return ofd.FileName;
        }
        #endregion

        #region txtを開き、ファイル名をreturn
        public static string OpenTxtFileDialog()
        {
            var ofd = new Microsoft.Win32.OpenFileDialog()
            {
               
                Filter = "10000bytefiles(*.*)|*.*|All files(*.*)|*.*",
                FilterIndex = 1
            };

            if (ofd.ShowDialog() != true)
                return "";
            return ofd.FileName;
        }

        #endregion

        #region txtをロード
        public static DataTable ImportTxtFile()
        {

            var filepath = OpenTxtFileDialog();
            if (filepath != "")
            {
                DataTable dt = new DataTable();
                try
                {
                    IList<List<string>> list = new List<List<string>>();
                    list.Add(TableTitle());
                    using (FileStream fsr = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                    {
                        using (StreamReader sr = new StreamReader(fsr, Encoding.Default))
                        {
                            while (sr.Peek() >= 0)
                            {
                                string temp = sr.ReadLine();
                                Console.WriteLine(temp.Length);
                                list.Add(StringToList(temp));
                            }
                        }
                    }
                    dt = StringListToDataTable(list);
                }
                catch (Exception e)
                {

                    throw e;
                }

                return dt;
            }
            else
                return null;
        }
        #endregion

        #region excelをdatatableに変換
        public static DataTable ImportExcelFile()
        {
            DataTable dt = null;
            var filepath = OpenExcelFileDialog();
            if (filepath.EndsWith(".xlsx"))
            {
                dt = Read2007Excel(filepath);
            }
            else if (filepath.EndsWith(".xls"))
            {
                dt = Read2003Excel(filepath);
            }
            else if (filepath.EndsWith(".csv"))
            {
                dt = CSVFileHelper.OpenCSV(filepath);
            }
            else
            {
                return dt;
            }
            return dt;
        }

        private static DataTable Read2007Excel(string filepath)
        {
            DataTable dt = new DataTable();
            if (filepath != null)
            {
                XSSFWorkbook xssfworkbook = null;

                var sheet = xssfworkbook.GetSheetAt(0);
                IEnumerator rows = sheet.GetRowEnumerator();

                IRow headRow = sheet.GetRow(0);

                for (int i = headRow.FirstCellNum, len = headRow.LastCellNum; i < len; i++)
                {
                    dt.Columns.Add(headRow.Cells[i].StringCellValue);
                }


                //for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
                //{
                //    dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
                //}
                rows.MoveNext();
                while (rows.MoveNext())
                {
                    XSSFRow row = (XSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();
                    if (row.GetCell(0).ToString().Equals(""))
                    {
                        continue;
                    }
                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);

                        if (cell == null)
                        {
                            dr[i] = "";
                        }
                        else
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    dr[i] = cell.DateCellValue;
                                }
                                else
                                {
                                    dr[i] = cell.NumericCellValue;
                                }
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                            {
                                dr[i] = cell.BooleanCellValue;
                            }
                            else
                            {
                                dr[i] = cell.StringCellValue;
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        private static DataTable Read2003Excel(string filepath)
        {
            DataTable dt = new DataTable();
            if (filepath != null)
            {
                HSSFWorkbook hssfworkbook = null;

                var sheet = hssfworkbook.GetSheetAt(0);
                IEnumerator rows = sheet.GetRowEnumerator();

                IRow headRow = sheet.GetRow(0);

                for (int i = headRow.FirstCellNum, len = headRow.LastCellNum; i < len; i++)
                {
                    dt.Columns.Add(headRow.Cells[i].StringCellValue);
                }

                rows.MoveNext();
                while (rows.MoveNext())
                {
                    HSSFRow row = (HSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();
                    if ("".Equals(row.GetCell(0).ToString()))
                    {
                        continue;
                    }
                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);

                        if (cell == null)
                        {
                            dr[i] = "";
                        }
                        else
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                if (HSSFDateUtil.IsCellDateFormatted(cell))
                                {
                                    dr[i] = cell.DateCellValue;
                                }
                                else
                                {
                                    dr[i] = cell.NumericCellValue;
                                }
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                            {
                                dr[i] = cell.BooleanCellValue;
                            }
                            else
                            {
                                dr[i] = cell.StringCellValue;
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        #endregion

        #region list<>をdatatableに変換
        public static DataTable ListToDataTable<T>(IEnumerable<T> c)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (c.Count() > 0)
            {
                for (int i = 0; i < c.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo item in props)
                    {
                        object obj = item.GetValue(c.ElementAt(i), null);
                        tempList.Add(obj);
                    }
                    dt.LoadDataRow(tempList.ToArray(), true);
                }
            }
            return dt;
        }
        #endregion

        #region list<List<string>>をdatatableに変換
        public static DataTable StringListToDataTable(IList<List<string>> list)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("testcase", Type.GetType("System.String")));
            for (int i = 0; i < list[0].Count; i++)
            {
                dt.Columns.Add(new DataColumn(list[0][i], Type.GetType("System.String")));
            }

            for (int i = 1; i < list.Count; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < list[0].Count; j++)
                {
                    dr[list[0][j]] = list[i][j];
                }
                dt.Rows.Add(dr);
            }


            return dt;
        }
        #endregion

        #region datatable を list<>に変換
        public static List<T> DataTableToList<T>(DataTable dataTable)
        {
            List<T> list = null;
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                list = new List<T>();

                Type t = typeof(T);

                PropertyInfo[] pinfo = t.GetProperties();

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    T item;
                    object objInstance = Activator.CreateInstance(t, true);

                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        string colName = dataTable.Columns[j].ColumnName;

                        for (int k = 0; k < pinfo.Length; k++)
                        {
                            if (colName.Equals(pinfo[k].Name, StringComparison.OrdinalIgnoreCase))
                            {
                                object defaultvalue = null;
                                object[] proAttributes = pinfo[k].GetCustomAttributes(typeof(DataValueDefaultAttribute), false);
                                if (proAttributes.Length > 0)
                                {
                                    DataValueDefaultAttribute loDefectTrack = (DataValueDefaultAttribute)proAttributes[0];

                                    defaultvalue = loDefectTrack.Value;
                                }

                                pinfo[k].SetValue(objInstance, Map(pinfo[k].PropertyType.ToString(), dataTable.Rows[i][colName].ToString(), defaultvalue), null);
                            }
                        }
                    }
                    item = (T)objInstance;
                    list.Add(item);
                }
            }
            return list;
        }

        private static object Map(string enType, string dbValue, object defaultvalue)
        {
            switch (enType.ToLower().Split('.')[1])
            {
                case "boolean":
                    try
                    {
                        return Convert.ToBoolean(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToBoolean(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "sbyte":
                    try
                    {
                        return Convert.ToSByte(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToSByte(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "int16":
                    try
                    {
                        return Convert.ToInt16(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToInt16(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "int32":
                    try
                    {
                        return Convert.ToInt32(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToInt32(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "int64":
                    try
                    {
                        return Convert.ToInt64(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToInt64(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "uint16":
                    try
                    {
                        return Convert.ToUInt16(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToUInt16(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "uint32":
                    try
                    {
                        return Convert.ToUInt32(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToUInt32(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "uint64":
                    try
                    {
                        return Convert.ToUInt64(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToUInt64(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "char":
                    try
                    {
                        return Convert.ToChar(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToChar(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "string":
                    try
                    {
                        return Convert.ToString(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToString(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "byte":
                    try
                    {
                        return Convert.ToByte(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToByte(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "single":
                    try
                    {
                        return Convert.ToSingle(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToSingle(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "double":
                    try
                    {
                        return Convert.ToDouble(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToDouble(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "decimal":
                    try
                    {
                        return Convert.ToDecimal(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToDecimal(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                case "datetime":
                    try
                    {
                        return Convert.ToDateTime(dbValue);
                    }
                    catch
                    {
                        if (defaultvalue != null)
                        {
                            return Convert.ToDateTime(defaultvalue);
                        }
                        else { throw new Exception("引数変換エラー"); }
                    }
                default:
                    return Convert.ToString(dbValue);
            }
        }
        #endregion

        #region List<>をローカルに
        public static bool SaveListToTxt<T>(IList<T> list, string path)
        {
            IoHelper.DeleteFile(path);
            using (FileStream fsw = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {

                using (StreamWriter sw = new StreamWriter(fsw))
                {
                    sw.WriteLine(JsonConvert.SerializeObject(list));
                }
            }
            if (IoHelper.Exists(path))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region datatablをローカルに
        public static bool SaveDataTableToTxt(DataTable dt, string path)
        {
            IoHelper.DeleteFile(path);
            using (FileStream fsw = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {

                using (StreamWriter sw = new StreamWriter(fsw))
                {
                    sw.WriteLine(JsonConvert.SerializeObject(dt));
                }
            }
            if (IoHelper.Exists(path))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region ローカルdatatableを回復
        public static DataTable TxtToDataTable(string path)
        {
            DataTable dt;
            using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
            {
                string line = sr.ReadLine();
                dt = JsonConvert.DeserializeObject<DataTable>(line);
            }
            return dt;
        }
        #endregion

        #region list<>をexcelに
        public static bool WriteExcel<T>(IList<T> list)
        {
            var filepath = SaveExcelFileDialog();
            if (filepath == "")
                return false;

            var dt = ListToDataTable<T>(list);
            IWorkbook book = null;

            if (filepath.EndsWith(".xlsx"))
            {
                book = new XSSFWorkbook();
            }
            else if (filepath.EndsWith(".xls"))
            {
                book = new HSSFWorkbook();
            }
            else if (filepath.EndsWith(".csv"))
            {
                return CSVFileHelper.SaveCSV(dt, filepath);
            }
            else
            {
                return false;
            }

            if (!string.IsNullOrEmpty(filepath) && null != dt && dt.Rows.Count > 0)
            {
                ISheet sheet = book.CreateSheet("Sheet1");
                IRow row = sheet.CreateRow(0);

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    NPOI.SS.UserModel.IRow row2 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row2.CreateCell(j).SetCellValue(Convert.ToString(dt.Rows[i][j]));

                    }
                }

                //设置格子宽高
                sheet.DefaultColumnWidth = 15;
                sheet.DefaultRowHeightInPoints = 20;

                //一般字体
                IFont fontFore = book.CreateFont(); //创建一个字体样式对象
                fontFore.FontName = "Meiryo UI"; //和excel里面的字体对应
                fontFore.FontHeightInPoints = 12;//字体大小

                //pass字体
                IFont fontPass = book.CreateFont(); //创建一个字体样式对象
                fontPass.FontName = "Meiryo UI"; //和excel里面的字体对应
                fontPass.FontHeightInPoints = 12;//字体大小
                fontPass.Color = 57;

                //fail字体
                IFont fontFail = book.CreateFont(); //创建一个字体样式对象
                fontFail.FontName = "Meiryo UI"; //和excel里面的字体对应
                fontFail.FontHeightInPoints = 12;//字体大小
                fontFail.Color = 10;

                //titel字体
                IFont fontTitle = book.CreateFont(); //创建一个字体样式对象
                fontTitle.FontName = "Meiryo UI"; //和excel里面的字体对应
                fontTitle.FontHeightInPoints = 14;//字体大小
                fontTitle.Color = 53;

                //subtitle字体
                IFont subTitleFont = book.CreateFont();
                subTitleFont.FontName = "Meiryo UI"; //和excel里面的字体对应
                subTitleFont.FontHeightInPoints = 13;//字体大小
                subTitleFont.Color = 18;

                //passed样式
                ICellStyle stylePass = book.CreateCellStyle();//创建样式对象
                stylePass.SetFont(fontPass); //将字体样式赋给样式对象
                stylePass.FillForegroundColor = 42;
                stylePass.FillPattern = FillPattern.SolidForeground;
                //设置单元格上下左右边框线  
                stylePass.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                stylePass.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                stylePass.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                stylePass.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                //文字水平和垂直对齐方式  
                stylePass.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                stylePass.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                //failed样式
                ICellStyle styleFail = book.CreateCellStyle();//创建样式对象
                styleFail.SetFont(fontFail); //将字体样式赋给样式对象
                styleFail.FillForegroundColor = 45;
                styleFail.FillPattern = FillPattern.SolidForeground;
                //设置单元格上下左右边框线  
                styleFail.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                styleFail.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                styleFail.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                styleFail.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                //文字水平和垂直对齐方式  
                styleFail.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                styleFail.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                //一般样式
                ICellStyle stylefore = book.CreateCellStyle();//创建样式对象
                stylefore.SetFont(fontFore); //将字体样式赋给样式对象
                //设置单元格上下左右边框线  
                stylefore.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                stylefore.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                stylefore.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                stylefore.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                //文字水平和垂直对齐方式  
                stylefore.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                stylefore.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                //title样式
                ICellStyle styleTitle = book.CreateCellStyle();//创建样式对象
                styleTitle.SetFont(fontTitle); //将字体样式赋给样式对象
                styleTitle.FillForegroundColor = 43;
                styleTitle.FillPattern = FillPattern.SolidForeground;
                //设置单元格上下左右边框线  
                styleTitle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                styleTitle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                styleTitle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                styleTitle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                //文字水平和垂直对齐方式  
                styleTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                styleTitle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                //subtitle样式
                ICellStyle styleSubTitle = book.CreateCellStyle();//创建样式对象
                styleSubTitle.SetFont(subTitleFont); //将字体样式赋给样式对象
                styleSubTitle.FillForegroundColor = 15;
                styleSubTitle.FillPattern = FillPattern.SolidForeground;
                //设置单元格上下左右边框线  
                styleSubTitle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                styleSubTitle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                styleSubTitle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                styleSubTitle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                //文字水平和垂直对齐方式  
                styleSubTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                styleSubTitle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;


                IRow firstRow = sheet.GetRow(0);//第一行
                int cellCount = firstRow.LastCellNum;//列数
                int rowCount = sheet.LastRowNum;//总行数
                string value;
                ICell cell;
                for (int i = 0; i <= rowCount; i++)
                {
                    for (int j = 0; j < cellCount; j++)
                    {
                        cell = sheet.GetRow(i).GetCell(j);
                        value = cell.ToString();
                        if (value.Equals("EXPECTED VALUE") || value.Equals("ACTUAL VALUE") || value.Equals("RESULT") || value.Equals("ETL TOOL"))
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(i, i, 0, cellCount - 1));
                            cell.CellStyle = styleSubTitle;
                            continue;
                        }
                        if (i == 0 || i == 1 || j == 0)
                        {
                            cell.CellStyle = styleTitle;
                            continue;
                        }
                        if (value.Equals("Passed"))
                        {
                            cell.CellStyle = stylePass;
                        }
                        else if (value.Equals("Failed"))
                        {
                            cell.CellStyle = styleFail;
                        }
                        else
                        {
                            cell.CellStyle = stylefore;
                        }
                    }
                }


                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    book.Write(ms);
                    using (FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                    book = null;
                }

            }
            return true;
        }
        #endregion

        #region  datatableをexcelに
        public static bool DataTableToExcel(DataTable dt, string path, string sheetName)
        {
            XSSFWorkbook book = null;
            using (FileStream fs = File.Open(path, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite))
            {
                book = new XSSFWorkbook(fs);
                if (!string.IsNullOrEmpty(path) && null != dt && dt.Rows.Count > 0)
                {

                    try
                    {
                        if (book.GetSheet(sheetName).SheetName == sheetName)
                        {
                            return false;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    //新增sheet
                    ISheet sheet = book.CreateSheet(sheetName);
                    IRow rowTitle = sheet.CreateRow(0);


                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        rowTitle.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        IRow row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            row.CreateCell(j).SetCellValue(Convert.ToString(dt.Rows[i][j]));

                        }
                    }

                    



                    //设置格子宽高
                    sheet.DefaultColumnWidth = 15;
                    sheet.DefaultRowHeightInPoints = 20;

                    //一般字体
                    IFont fontFore = book.CreateFont(); //创建一个字体样式对象
                    fontFore.FontName = "Meiryo UI"; //和excel里面的字体对应
                    fontFore.FontHeightInPoints = 12;//字体大小

                    //pass字体
                    IFont fontPass = book.CreateFont(); //创建一个字体样式对象
                    fontPass.FontName = "Meiryo UI"; //和excel里面的字体对应
                    fontPass.FontHeightInPoints = 12;//字体大小
                    fontPass.Color = 57;

                    //fail字体
                    IFont fontFail = book.CreateFont(); //创建一个字体样式对象
                    fontFail.FontName = "Meiryo UI"; //和excel里面的字体对应
                    fontFail.FontHeightInPoints = 12;//字体大小
                    fontFail.Color = 10;

                    //titel字体
                    IFont fontTitle = book.CreateFont(); //创建一个字体样式对象
                    fontTitle.FontName = "Meiryo UI"; //和excel里面的字体对应
                    fontTitle.FontHeightInPoints = 14;//字体大小
                    fontTitle.Color = 53;

                    //subtitle字体
                    IFont subTitleFont = book.CreateFont();
                    subTitleFont.FontName = "Meiryo UI"; //和excel里面的字体对应
                    subTitleFont.FontHeightInPoints = 13;//字体大小
                    subTitleFont.Color = 18;

                    //passed样式
                    ICellStyle stylePass = book.CreateCellStyle();//创建样式对象
                    stylePass.SetFont(fontPass); //将字体样式赋给样式对象
                    stylePass.FillForegroundColor = 42;
                    stylePass.FillPattern = FillPattern.SolidForeground;
                    //设置单元格上下左右边框线  
                    stylePass.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    stylePass.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    stylePass.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    stylePass.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //文字水平和垂直对齐方式  
                    stylePass.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    stylePass.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    //failed样式
                    ICellStyle styleFail = book.CreateCellStyle();//创建样式对象
                    styleFail.SetFont(fontFail); //将字体样式赋给样式对象
                    styleFail.FillForegroundColor = 45;
                    styleFail.FillPattern = FillPattern.SolidForeground;
                    //设置单元格上下左右边框线  
                    styleFail.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleFail.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleFail.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleFail.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //文字水平和垂直对齐方式  
                    styleFail.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    styleFail.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    //一般样式
                    ICellStyle stylefore = book.CreateCellStyle();//创建样式对象
                    stylefore.SetFont(fontFore); //将字体样式赋给样式对象
                    //设置单元格上下左右边框线  
                    stylefore.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    stylefore.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    stylefore.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    stylefore.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //文字水平和垂直对齐方式  
                    stylefore.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                    stylefore.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    //title样式
                    ICellStyle styleTitle = book.CreateCellStyle();//创建样式对象
                    styleTitle.SetFont(fontTitle); //将字体样式赋给样式对象
                    styleTitle.FillForegroundColor = 43;
                    styleTitle.FillPattern = FillPattern.SolidForeground;
                    //设置单元格上下左右边框线  
                    styleTitle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleTitle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleTitle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleTitle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //文字水平和垂直对齐方式  
                    styleTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    styleTitle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

                    //subtitle样式
                    ICellStyle styleSubTitle = book.CreateCellStyle();//创建样式对象
                    styleSubTitle.SetFont(subTitleFont); //将字体样式赋给样式对象
                    styleSubTitle.FillForegroundColor = 15;
                    styleSubTitle.FillPattern = FillPattern.SolidForeground;
                    //设置单元格上下左右边框线  
                    styleSubTitle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleSubTitle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleSubTitle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    styleSubTitle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //文字水平和垂直对齐方式  
                    styleSubTitle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    styleSubTitle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;


                    IRow firstRow = sheet.GetRow(0);//第一行
                    int cellCount = firstRow.LastCellNum;//列数
                    int rowCount = sheet.LastRowNum;//总行数
                    string value;
                    ICell cell;
                    for (int i = 0; i <= rowCount; i++)
                    {
                        for (int j = 0; j < cellCount; j++)
                        {
                            cell = sheet.GetRow(i).GetCell(j);
                            value = cell.ToString();
                            //把样式赋给单元格
                            if (value.Equals("EXPECTED VALUE") || value.Equals("ACTUAL VALUE") || value.Equals("RESULT") || value.Equals("ETL TOOL"))
                            {
                                //合并副标题整行
                                sheet.AddMergedRegion(new CellRangeAddress(i, i, 0, cellCount - 1));
                                cell.CellStyle = styleSubTitle;
                                continue;
                            }
                            if (i == 0 || j == 0)
                            {
                                cell.CellStyle = styleTitle;
                                continue;
                            }
                            if (value.Equals("Passed"))
                            {
                                cell.CellStyle = stylePass;
                            }
                            else if (value.Equals("Failed"))
                            {
                                cell.CellStyle = styleFail;
                            }
                            else
                            {
                                cell.CellStyle = stylefore;
                            }
                        }
                    }
                }
                //保存操作
                using (FileStream fileStream = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    book.Write(fileStream);
                    fileStream.Close();
                }

                return true;
            }
        }
        #endregion

        #region 10000byteファイルをListに変換
        public static List<string> StringToList(string str)
        {
            List<string> list = new List<string>();
            list.Add(str.Substring(IndexMinusOne(1), 10));
            list.Add(str.Substring(IndexMinusOne(11), 1));
            list.Add(str.Substring(IndexMinusOne(12), 1));
            list.Add(str.Substring(IndexMinusOne(13), 1));
            list.Add(str.Substring(IndexMinusOne(14), 1));
            list.Add(str.Substring(IndexMinusOne(15), 8));
            list.Add(str.Substring(IndexMinusOne(23), 10));
            list.Add(str.Substring(IndexMinusOne(33), 11));
            list.Add(str.Substring(IndexMinusOne(44), 2));
            list.Add(str.Substring(IndexMinusOne(46), 11));
            list.Add(str.Substring(IndexMinusOne(57), 2));
            list.Add(str.Substring(IndexMinusOne(59), 2));
            list.Add(str.Substring(IndexMinusOne(61), 2));
            list.Add(str.Substring(IndexMinusOne(63), 2));
            list.Add(str.Substring(IndexMinusOne(65), 2));
            list.Add(str.Substring(IndexMinusOne(67), 8));
            list.Add(str.Substring(IndexMinusOne(75), 8));
            list.Add(str.Substring(IndexMinusOne(83), 8));
            list.Add(str.Substring(IndexMinusOne(91), 8));
            list.Add(str.Substring(IndexMinusOne(99), 6));
            list.Add(str.Substring(IndexMinusOne(105), 1));
            list.Add(str.Substring(IndexMinusOne(106), 30));
            list.Add(str.Substring(IndexMinusOne(136), 60));
            list.Add(str.Substring(IndexMinusOne(196), 8));
            list.Add(str.Substring(IndexMinusOne(204), 2));
            list.Add(str.Substring(IndexMinusOne(206), 16));
            list.Add(str.Substring(IndexMinusOne(222), 7));
            list.Add(str.Substring(IndexMinusOne(229), 30));
            list.Add(str.Substring(IndexMinusOne(259), 30));
            list.Add(str.Substring(IndexMinusOne(289), 30));
            list.Add(str.Substring(IndexMinusOne(319), 40));
            list.Add(str.Substring(IndexMinusOne(359), 6));
            list.Add(str.Substring(IndexMinusOne(365), 6));
            list.Add(str.Substring(IndexMinusOne(371), 1));
            list.Add(str.Substring(IndexMinusOne(372), 1));
            list.Add(str.Substring(IndexMinusOne(373), 6));
            list.Add(str.Substring(IndexMinusOne(379), 6));
            list.Add(str.Substring(IndexMinusOne(385), 3));
            list.Add(str.Substring(IndexMinusOne(388), 3));
            list.Add(str.Substring(IndexMinusOne(391), 5));
            list.Add(str.Substring(IndexMinusOne(396), 5));
            list.Add(str.Substring(IndexMinusOne(401), 1));
            list.Add(str.Substring(IndexMinusOne(402), 1));
            list.Add(str.Substring(IndexMinusOne(403), 5));
            list.Add(str.Substring(IndexMinusOne(408), 1));
            list.Add(str.Substring(IndexMinusOne(409), 1));
            list.Add(str.Substring(IndexMinusOne(410), 11));
            list.Add(str.Substring(IndexMinusOne(421), 1));
            list.Add(str.Substring(IndexMinusOne(422), 2));
            list.Add(str.Substring(IndexMinusOne(424), 1));
            list.Add(str.Substring(IndexMinusOne(425), 8));
            list.Add(str.Substring(IndexMinusOne(433), 1));
            list.Add(str.Substring(IndexMinusOne(434), 1));
            list.Add(str.Substring(IndexMinusOne(435), 1));
            list.Add(str.Substring(IndexMinusOne(436), 1));
            list.Add(str.Substring(IndexMinusOne(437), 1));
            list.Add(str.Substring(IndexMinusOne(438), 1));
            list.Add(str.Substring(IndexMinusOne(439), 1));
            list.Add(str.Substring(IndexMinusOne(440), 30));
            list.Add(str.Substring(IndexMinusOne(470), 60));
            list.Add(str.Substring(IndexMinusOne(530), 8));
            list.Add(str.Substring(IndexMinusOne(538), 2));
            list.Add(str.Substring(IndexMinusOne(540), 6));
            list.Add(str.Substring(IndexMinusOne(546), 1));
            list.Add(str.Substring(IndexMinusOne(547), 1));
            list.Add(str.Substring(IndexMinusOne(548), 30));
            list.Add(str.Substring(IndexMinusOne(578), 60));
            list.Add(str.Substring(IndexMinusOne(638), 2));
            list.Add(str.Substring(IndexMinusOne(640), 7));
            list.Add(str.Substring(IndexMinusOne(647), 1111));
            list.Add(str.Substring(IndexMinusOne(1758), 110));
            list.Add(str.Substring(IndexMinusOne(1868), 10));
            list.Add(str.Substring(IndexMinusOne(1878), 1));
            list.Add(str.Substring(IndexMinusOne(1879), 3));
            list.Add(str.Substring(IndexMinusOne(1882), 2));
            list.Add(str.Substring(IndexMinusOne(1884), 2));
            list.Add(str.Substring(IndexMinusOne(1886), 2));
            list.Add(str.Substring(IndexMinusOne(1888), 2));
            list.Add(str.Substring(IndexMinusOne(1890), 2));
            list.Add(str.Substring(IndexMinusOne(1892), 11));
            list.Add(str.Substring(IndexMinusOne(1903), 3));
            list.Add(str.Substring(IndexMinusOne(1906), 2));
            list.Add(str.Substring(IndexMinusOne(1908), 3));
            list.Add(str.Substring(IndexMinusOne(1911), 2));
            list.Add(str.Substring(IndexMinusOne(1913), 2));
            list.Add(str.Substring(IndexMinusOne(1915), 2));
            list.Add(str.Substring(IndexMinusOne(1917), 2));
            list.Add(str.Substring(IndexMinusOne(1919), 2));
            list.Add(str.Substring(IndexMinusOne(1921), 11));
            list.Add(str.Substring(IndexMinusOne(1932), 3));
            list.Add(str.Substring(IndexMinusOne(1935), 2));
            list.Add(str.Substring(IndexMinusOne(1937), 3));
            list.Add(str.Substring(IndexMinusOne(1940), 2));
            list.Add(str.Substring(IndexMinusOne(1942), 2));
            list.Add(str.Substring(IndexMinusOne(1944), 2));
            list.Add(str.Substring(IndexMinusOne(1946), 2));
            list.Add(str.Substring(IndexMinusOne(1948), 2));
            list.Add(str.Substring(IndexMinusOne(1950), 11));
            list.Add(str.Substring(IndexMinusOne(1961), 3));
            list.Add(str.Substring(IndexMinusOne(1964), 2));
            list.Add(str.Substring(IndexMinusOne(1966), 493));
            list.Add(str.Substring(IndexMinusOne(2459), 1));
            list.Add(str.Substring(IndexMinusOne(2460), 2));
            list.Add(str.Substring(IndexMinusOne(2462), 2));
            list.Add(str.Substring(IndexMinusOne(2464), 11));
            list.Add(str.Substring(IndexMinusOne(2475), 11));
            list.Add(str.Substring(IndexMinusOne(2486), 11));
            list.Add(str.Substring(IndexMinusOne(2497), 1));
            list.Add(str.Substring(IndexMinusOne(2498), 1));
            list.Add(str.Substring(IndexMinusOne(2499), 1));
            list.Add(str.Substring(IndexMinusOne(2500), 1));
            list.Add(str.Substring(IndexMinusOne(2501), 11));
            list.Add(str.Substring(IndexMinusOne(2512), 1));
            list.Add(str.Substring(IndexMinusOne(2513), 2));
            list.Add(str.Substring(IndexMinusOne(2515), 1));
            list.Add(str.Substring(IndexMinusOne(2516), 11));
            list.Add(str.Substring(IndexMinusOne(2527), 11));
            list.Add(str.Substring(IndexMinusOne(2538), 1));
            list.Add(str.Substring(IndexMinusOne(2539), 30));
            list.Add(str.Substring(IndexMinusOne(2569), 60));
            list.Add(str.Substring(IndexMinusOne(2629), 8));
            list.Add(str.Substring(IndexMinusOne(2637), 1));
            list.Add(str.Substring(IndexMinusOne(2638), 2));
            list.Add(str.Substring(IndexMinusOne(2640), 3));
            list.Add(str.Substring(IndexMinusOne(2643), 1));
            list.Add(str.Substring(IndexMinusOne(2644), 11));
            list.Add(str.Substring(IndexMinusOne(2655), 10));
            list.Add(str.Substring(IndexMinusOne(2665), 4));
            list.Add(str.Substring(IndexMinusOne(2669), 30));
            list.Add(str.Substring(IndexMinusOne(2699), 40));
            list.Add(str.Substring(IndexMinusOne(2739), 3));
            list.Add(str.Substring(IndexMinusOne(2742), 30));
            list.Add(str.Substring(IndexMinusOne(2772), 40));
            list.Add(str.Substring(IndexMinusOne(2812), 1));
            list.Add(str.Substring(IndexMinusOne(2813), 7));
            list.Add(str.Substring(IndexMinusOne(2820), 28));
            list.Add(str.Substring(IndexMinusOne(2848), 28));
            list.Add(str.Substring(IndexMinusOne(2876), 5));
            list.Add(str.Substring(IndexMinusOne(2881), 3));
            list.Add(str.Substring(IndexMinusOne(2884), 2));
            list.Add(str.Substring(IndexMinusOne(2886), 10));
            list.Add(str.Substring(IndexMinusOne(2896), 7));
            list.Add(str.Substring(IndexMinusOne(2903), 3));
            list.Add(str.Substring(IndexMinusOne(2906), 5));
            list.Add(str.Substring(IndexMinusOne(2911), 1));
            list.Add(str.Substring(IndexMinusOne(2912), 6));
            list.Add(str.Substring(IndexMinusOne(2918), 2));
            list.Add(str.Substring(IndexMinusOne(2920), 28));
            list.Add(str.Substring(IndexMinusOne(2948), 16));
            list.Add(str.Substring(IndexMinusOne(2964), 7));
            list.Add(str.Substring(IndexMinusOne(2971), 30));
            list.Add(str.Substring(IndexMinusOne(3001), 30));
            list.Add(str.Substring(IndexMinusOne(3031), 30));
            list.Add(str.Substring(IndexMinusOne(3061), 1));
            list.Add(str.Substring(IndexMinusOne(3062), 8));
            list.Add(str.Substring(IndexMinusOne(3070), 1));
            list.Add(str.Substring(IndexMinusOne(3071), 11));
            list.Add(str.Substring(IndexMinusOne(3082), 7));
            list.Add(str.Substring(IndexMinusOne(3089), 1));
            list.Add(str.Substring(IndexMinusOne(3090), 2));
            list.Add(str.Substring(IndexMinusOne(3092), 11));
            list.Add(str.Substring(IndexMinusOne(3103), 11));
            list.Add(str.Substring(IndexMinusOne(3114), 1));
            list.Add(str.Substring(IndexMinusOne(3115), 16));
            list.Add(str.Substring(IndexMinusOne(3131), 6));
            list.Add(str.Substring(IndexMinusOne(3137), 30));
            list.Add(str.Substring(IndexMinusOne(3167), 7));
            list.Add(str.Substring(IndexMinusOne(3174), 7));
            list.Add(str.Substring(IndexMinusOne(3181), 11));
            list.Add(str.Substring(IndexMinusOne(3192), 11));
            list.Add(str.Substring(IndexMinusOne(3203), 8));
            list.Add(str.Substring(IndexMinusOne(3211), 5));
            list.Add(str.Substring(IndexMinusOne(3216), 10));
            list.Add(str.Substring(IndexMinusOne(3226), 30));
            list.Add(str.Substring(IndexMinusOne(3256), 1));
            list.Add(str.Substring(IndexMinusOne(3257), 1));
            list.Add(str.Substring(IndexMinusOne(3258), 4));
            list.Add(str.Substring(IndexMinusOne(3262), 1));
            list.Add(str.Substring(IndexMinusOne(3263), 2));
            list.Add(str.Substring(IndexMinusOne(3265), 36));
            list.Add(str.Substring(IndexMinusOne(3301), 8));
            list.Add(str.Substring(IndexMinusOne(3309), 1));
            list.Add(str.Substring(IndexMinusOne(3310), 1));
            list.Add(str.Substring(IndexMinusOne(3311), 1));
            list.Add(str.Substring(IndexMinusOne(3312), 1));
            list.Add(str.Substring(IndexMinusOne(3313), 1));
            list.Add(str.Substring(IndexMinusOne(3314), 1));
            list.Add(str.Substring(IndexMinusOne(3315), 2));
            list.Add(str.Substring(IndexMinusOne(3317), 11));
            list.Add(str.Substring(IndexMinusOne(3328), 8));
            list.Add(str.Substring(IndexMinusOne(3336), 3));
            list.Add(str.Substring(IndexMinusOne(3339), 1));
            list.Add(str.Substring(IndexMinusOne(3340), 1));
            list.Add(str.Substring(IndexMinusOne(3341), 4));
            list.Add(str.Substring(IndexMinusOne(3345), 4));
            list.Add(str.Substring(IndexMinusOne(3349), 2));
            list.Add(str.Substring(IndexMinusOne(3351), 1));
            list.Add(str.Substring(IndexMinusOne(3352), 1));
            list.Add(str.Substring(IndexMinusOne(3353), 6));
            list.Add(str.Substring(IndexMinusOne(3359), 10));
            list.Add(str.Substring(IndexMinusOne(3369), 10));
            list.Add(str.Substring(IndexMinusOne(3379), 1));
            list.Add(str.Substring(IndexMinusOne(3380), 1));
            list.Add(str.Substring(IndexMinusOne(3381), 1));
            list.Add(str.Substring(IndexMinusOne(3382), 8));
            list.Add(str.Substring(IndexMinusOne(3390), 1));
            list.Add(str.Substring(IndexMinusOne(3391), 1));
            list.Add(str.Substring(IndexMinusOne(3392), 1));
            list.Add(str.Substring(IndexMinusOne(3393), 5));
            list.Add(str.Substring(IndexMinusOne(3398), 1));
            list.Add(str.Substring(IndexMinusOne(3399), 5));
            list.Add(str.Substring(IndexMinusOne(3404), 1));
            list.Add(str.Substring(IndexMinusOne(3405), 1));
            list.Add(str.Substring(IndexMinusOne(3406), 1));
            list.Add(str.Substring(IndexMinusOne(3407), 1));
            list.Add(str.Substring(IndexMinusOne(3408), 2));
            list.Add(str.Substring(IndexMinusOne(3410), 3));
            list.Add(str.Substring(IndexMinusOne(3413), 1));
            list.Add(str.Substring(IndexMinusOne(3414), 2));
            list.Add(str.Substring(IndexMinusOne(3416), 1));
            list.Add(str.Substring(IndexMinusOne(3417), 7));
            list.Add(str.Substring(IndexMinusOne(3424), 7));
            list.Add(str.Substring(IndexMinusOne(3431), 7));
            list.Add(str.Substring(IndexMinusOne(3438), 6));
            list.Add(str.Substring(IndexMinusOne(3444), 8));
            list.Add(str.Substring(IndexMinusOne(3452), 5));
            list.Add(str.Substring(IndexMinusOne(3457), 5));
            list.Add(str.Substring(IndexMinusOne(3462), 1));
            list.Add(str.Substring(IndexMinusOne(3463), 1));
            list.Add(str.Substring(IndexMinusOne(3464), 2));
            list.Add(str.Substring(IndexMinusOne(3466), 8));
            list.Add(str.Substring(IndexMinusOne(3474), 1));
            list.Add(str.Substring(IndexMinusOne(3475), 1));
            list.Add(str.Substring(IndexMinusOne(3476), 1));
            list.Add(str.Substring(IndexMinusOne(3477), 6));
            list.Add(str.Substring(IndexMinusOne(3483), 1));
            list.Add(str.Substring(IndexMinusOne(3484), 1));
            list.Add(str.Substring(IndexMinusOne(3485), 1));
            list.Add(str.Substring(IndexMinusOne(3486), 1));
            list.Add(str.Substring(IndexMinusOne(3487), 1));
            list.Add(str.Substring(IndexMinusOne(3488), 6));
            list.Add(str.Substring(IndexMinusOne(3494), 6));
            list.Add(str.Substring(IndexMinusOne(3500), 1));
            list.Add(str.Substring(IndexMinusOne(3501), 1));
            list.Add(str.Substring(IndexMinusOne(3502), 5));
            list.Add(str.Substring(IndexMinusOne(3507), 1));
            list.Add(str.Substring(IndexMinusOne(3508), 2));
            list.Add(str.Substring(IndexMinusOne(3510), 6));
            list.Add(str.Substring(IndexMinusOne(3516), 6));
            list.Add(str.Substring(IndexMinusOne(3522), 1));
            list.Add(str.Substring(IndexMinusOne(3523), 8));
            list.Add(str.Substring(IndexMinusOne(3531), 1));
            list.Add(str.Substring(IndexMinusOne(3532), 6));
            list.Add(str.Substring(IndexMinusOne(3538), 7));
            list.Add(str.Substring(IndexMinusOne(3545), 5));
            list.Add(str.Substring(IndexMinusOne(3550), 7));
            list.Add(str.Substring(IndexMinusOne(3557), 7));
            list.Add(str.Substring(IndexMinusOne(3564), 1));
            list.Add(str.Substring(IndexMinusOne(3565), 1));
            list.Add(str.Substring(IndexMinusOne(3566), 1));
            list.Add(str.Substring(IndexMinusOne(3567), 1));
            list.Add(str.Substring(IndexMinusOne(3568), 5));
            list.Add(str.Substring(IndexMinusOne(3573), 5));
            list.Add(str.Substring(IndexMinusOne(3578), 5));
            list.Add(str.Substring(IndexMinusOne(3583), 1));
            list.Add(str.Substring(IndexMinusOne(3584), 5));
            list.Add(str.Substring(IndexMinusOne(3589), 1));
            list.Add(str.Substring(IndexMinusOne(3590), 5));
            list.Add(str.Substring(IndexMinusOne(3595), 1));
            list.Add(str.Substring(IndexMinusOne(3596), 1));
            list.Add(str.Substring(IndexMinusOne(3597), 1));
            list.Add(str.Substring(IndexMinusOne(3598), 3));
            list.Add(str.Substring(IndexMinusOne(3601), 1));
            list.Add(str.Substring(IndexMinusOne(3602), 3));
            list.Add(str.Substring(IndexMinusOne(3605), 1));
            list.Add(str.Substring(IndexMinusOne(3606), 3));
            list.Add(str.Substring(IndexMinusOne(3609), 1));
            list.Add(str.Substring(IndexMinusOne(3610), 3));
            list.Add(str.Substring(IndexMinusOne(3613), 1));
            list.Add(str.Substring(IndexMinusOne(3614), 3));
            list.Add(str.Substring(IndexMinusOne(3617), 6));
            list.Add(str.Substring(IndexMinusOne(3623), 6));
            list.Add(str.Substring(IndexMinusOne(3629), 1));
            list.Add(str.Substring(IndexMinusOne(3630), 6));
            list.Add(str.Substring(IndexMinusOne(3636), 6));
            list.Add(str.Substring(IndexMinusOne(3642), 6));
            list.Add(str.Substring(IndexMinusOne(3648), 1));
            list.Add(str.Substring(IndexMinusOne(3649), 1));
            list.Add(str.Substring(IndexMinusOne(3650), 1));
            list.Add(str.Substring(IndexMinusOne(3651), 5));
            list.Add(str.Substring(IndexMinusOne(3656), 1));
            list.Add(str.Substring(IndexMinusOne(3657), 1));
            list.Add(str.Substring(IndexMinusOne(3658), 1));
            list.Add(str.Substring(IndexMinusOne(3659), 6));
            list.Add(str.Substring(IndexMinusOne(3665), 1));
            list.Add(str.Substring(IndexMinusOne(3666), 1));
            list.Add(str.Substring(IndexMinusOne(3667), 30));
            list.Add(str.Substring(IndexMinusOne(3697), 60));
            list.Add(str.Substring(IndexMinusOne(3757), 8));
            list.Add(str.Substring(IndexMinusOne(3765), 1));
            list.Add(str.Substring(IndexMinusOne(3766), 30));
            list.Add(str.Substring(IndexMinusOne(3796), 60));
            list.Add(str.Substring(IndexMinusOne(3856), 8));
            list.Add(str.Substring(IndexMinusOne(3864), 1));
            list.Add(str.Substring(IndexMinusOne(3865), 30));
            list.Add(str.Substring(IndexMinusOne(3895), 60));
            list.Add(str.Substring(IndexMinusOne(3955), 8));
            list.Add(str.Substring(IndexMinusOne(3963), 1));
            list.Add(str.Substring(IndexMinusOne(3964), 30));
            list.Add(str.Substring(IndexMinusOne(3994), 60));
            list.Add(str.Substring(IndexMinusOne(4054), 8));
            list.Add(str.Substring(IndexMinusOne(4062), 1));
            list.Add(str.Substring(IndexMinusOne(4063), 60));
            list.Add(str.Substring(IndexMinusOne(4123), 60));
            list.Add(str.Substring(IndexMinusOne(4183), 60));
            list.Add(str.Substring(IndexMinusOne(4243), 8));
            list.Add(str.Substring(IndexMinusOne(4251), 6));
            list.Add(str.Substring(IndexMinusOne(4257), 2));
            list.Add(str.Substring(IndexMinusOne(4259), 11));
            list.Add(str.Substring(IndexMinusOne(4270), 11));
            list.Add(str.Substring(IndexMinusOne(4281), 1));
            list.Add(str.Substring(IndexMinusOne(4282), 30));
            list.Add(str.Substring(IndexMinusOne(4312), 60));
            list.Add(str.Substring(IndexMinusOne(4372), 8));
            list.Add(str.Substring(IndexMinusOne(4380), 1));
            list.Add(str.Substring(IndexMinusOne(4381), 1));
            list.Add(str.Substring(IndexMinusOne(4382), 30));
            list.Add(str.Substring(IndexMinusOne(4412), 60));
            list.Add(str.Substring(IndexMinusOne(4472), 8));
            list.Add(str.Substring(IndexMinusOne(4480), 1));
            list.Add(str.Substring(IndexMinusOne(4481), 10));
            list.Add(str.Substring(IndexMinusOne(4491), 10));
            list.Add(str.Substring(IndexMinusOne(4501), 10));
            list.Add(str.Substring(IndexMinusOne(4511), 3));
            list.Add(str.Substring(IndexMinusOne(4514), 2));
            list.Add(str.Substring(IndexMinusOne(4516), 1));
            list.Add(str.Substring(IndexMinusOne(4517), 1));
            list.Add(str.Substring(IndexMinusOne(4518), 7));
            list.Add(str.Substring(IndexMinusOne(4525), 1));
            list.Add(str.Substring(IndexMinusOne(4526), 8));
            list.Add(str.Substring(IndexMinusOne(4534), 1));
            list.Add(str.Substring(IndexMinusOne(4535), 10));
            list.Add(str.Substring(IndexMinusOne(4545), 1));
            list.Add(str.Substring(IndexMinusOne(4546), 435));
            list.Add(str.Substring(IndexMinusOne(4981), 2));
            list.Add(str.Substring(IndexMinusOne(4983), 15));
            list.Add(str.Substring(IndexMinusOne(4998), 1));
            list.Add(str.Substring(IndexMinusOne(4999), 11));
            list.Add(str.Substring(IndexMinusOne(5010), 11));
            list.Add(str.Substring(IndexMinusOne(5021), 1));
            list.Add(str.Substring(IndexMinusOne(5022), 126));
            list.Add(str.Substring(IndexMinusOne(5148), 3));
            list.Add(str.Substring(IndexMinusOne(5151), 60));
            list.Add(str.Substring(IndexMinusOne(5211), 2));
            list.Add(str.Substring(IndexMinusOne(5213), 8));
            list.Add(str.Substring(IndexMinusOne(5221), 30));
            list.Add(str.Substring(IndexMinusOne(5251), 30));
            list.Add(str.Substring(IndexMinusOne(5281), 30));
            list.Add(str.Substring(IndexMinusOne(5311), 70));
            list.Add(str.Substring(IndexMinusOne(5381), 1));
            list.Add(str.Substring(IndexMinusOne(5382), 1));
            list.Add(str.Substring(IndexMinusOne(5383), 1));
            list.Add(str.Substring(IndexMinusOne(5384), 1));
            list.Add(str.Substring(IndexMinusOne(5385), 8));
            list.Add(str.Substring(IndexMinusOne(5393), 8));
            list.Add(str.Substring(IndexMinusOne(5401), 4));
            list.Add(str.Substring(IndexMinusOne(5405), 5));
            list.Add(str.Substring(IndexMinusOne(5410), 12));
            list.Add(str.Substring(IndexMinusOne(5422), 13));
            list.Add(str.Substring(IndexMinusOne(5435), 7));
            list.Add(str.Substring(IndexMinusOne(5442), 4));
            list.Add(str.Substring(IndexMinusOne(5446), 5));
            list.Add(str.Substring(IndexMinusOne(5451), 12));
            list.Add(str.Substring(IndexMinusOne(5463), 13));
            list.Add(str.Substring(IndexMinusOne(5476), 7));
            list.Add(str.Substring(IndexMinusOne(5483), 7));
            list.Add(str.Substring(IndexMinusOne(5490), 3));
            list.Add(str.Substring(IndexMinusOne(5493), 2));
            list.Add(str.Substring(IndexMinusOne(5495), 2));
            list.Add(str.Substring(IndexMinusOne(5497), 2));
            list.Add(str.Substring(IndexMinusOne(5499), 2));
            list.Add(str.Substring(IndexMinusOne(5501), 2));
            list.Add(str.Substring(IndexMinusOne(5503), 11));
            list.Add(str.Substring(IndexMinusOne(5514), 3));
            list.Add(str.Substring(IndexMinusOne(5517), 2));
            list.Add(str.Substring(IndexMinusOne(5519), 2));
            list.Add(str.Substring(IndexMinusOne(5521), 11));
            list.Add(str.Substring(IndexMinusOne(5532), 8));
            list.Add(str.Substring(IndexMinusOne(5540), 11));
            list.Add(str.Substring(IndexMinusOne(5551), 11));
            list.Add(str.Substring(IndexMinusOne(5562), 11));
            list.Add(str.Substring(IndexMinusOne(5573), 8));
            list.Add(str.Substring(IndexMinusOne(5581), 1));
            list.Add(str.Substring(IndexMinusOne(5582), 8));
            list.Add(str.Substring(IndexMinusOne(5590), 10));
            list.Add(str.Substring(IndexMinusOne(5600), 11));
            list.Add(str.Substring(IndexMinusOne(5611), 3));
            list.Add(str.Substring(IndexMinusOne(5614), 2));
            list.Add(str.Substring(IndexMinusOne(5616), 3));
            list.Add(str.Substring(IndexMinusOne(5619), 2));
            list.Add(str.Substring(IndexMinusOne(5621), 2));
            list.Add(str.Substring(IndexMinusOne(5623), 2));
            list.Add(str.Substring(IndexMinusOne(5625), 2));
            list.Add(str.Substring(IndexMinusOne(5627), 2));
            list.Add(str.Substring(IndexMinusOne(5629), 11));
            list.Add(str.Substring(IndexMinusOne(5640), 3));
            list.Add(str.Substring(IndexMinusOne(5643), 2));
            list.Add(str.Substring(IndexMinusOne(5645), 2));
            list.Add(str.Substring(IndexMinusOne(5647), 11));
            list.Add(str.Substring(IndexMinusOne(5658), 8));
            list.Add(str.Substring(IndexMinusOne(5666), 11));
            list.Add(str.Substring(IndexMinusOne(5677), 11));
            list.Add(str.Substring(IndexMinusOne(5688), 11));
            list.Add(str.Substring(IndexMinusOne(5699), 8));
            list.Add(str.Substring(IndexMinusOne(5707), 1));
            list.Add(str.Substring(IndexMinusOne(5708), 8));
            list.Add(str.Substring(IndexMinusOne(5716), 10));
            list.Add(str.Substring(IndexMinusOne(5726), 11));
            list.Add(str.Substring(IndexMinusOne(5737), 3));
            list.Add(str.Substring(IndexMinusOne(5740), 2));
            list.Add(str.Substring(IndexMinusOne(5742), 3));
            list.Add(str.Substring(IndexMinusOne(5745), 2));
            list.Add(str.Substring(IndexMinusOne(5747), 2));
            list.Add(str.Substring(IndexMinusOne(5749), 2));
            list.Add(str.Substring(IndexMinusOne(5751), 2));
            list.Add(str.Substring(IndexMinusOne(5753), 2));
            list.Add(str.Substring(IndexMinusOne(5755), 11));
            list.Add(str.Substring(IndexMinusOne(5766), 3));
            list.Add(str.Substring(IndexMinusOne(5769), 2));
            list.Add(str.Substring(IndexMinusOne(5771), 2));
            list.Add(str.Substring(IndexMinusOne(5773), 11));
            list.Add(str.Substring(IndexMinusOne(5784), 8));
            list.Add(str.Substring(IndexMinusOne(5792), 11));
            list.Add(str.Substring(IndexMinusOne(5803), 11));
            list.Add(str.Substring(IndexMinusOne(5814), 11));
            list.Add(str.Substring(IndexMinusOne(5825), 8));
            list.Add(str.Substring(IndexMinusOne(5833), 1));
            list.Add(str.Substring(IndexMinusOne(5834), 8));
            list.Add(str.Substring(IndexMinusOne(5842), 10));
            list.Add(str.Substring(IndexMinusOne(5852), 11));
            list.Add(str.Substring(IndexMinusOne(5863), 3));
            list.Add(str.Substring(IndexMinusOne(5866), 2));
            list.Add(str.Substring(IndexMinusOne(5868), 3));
            list.Add(str.Substring(IndexMinusOne(5871), 2));
            list.Add(str.Substring(IndexMinusOne(5873), 2));
            list.Add(str.Substring(IndexMinusOne(5875), 2));
            list.Add(str.Substring(IndexMinusOne(5877), 2));
            list.Add(str.Substring(IndexMinusOne(5879), 2));
            list.Add(str.Substring(IndexMinusOne(5881), 11));
            list.Add(str.Substring(IndexMinusOne(5892), 3));
            list.Add(str.Substring(IndexMinusOne(5895), 2));
            list.Add(str.Substring(IndexMinusOne(5897), 2));
            list.Add(str.Substring(IndexMinusOne(5899), 11));
            list.Add(str.Substring(IndexMinusOne(5910), 8));
            list.Add(str.Substring(IndexMinusOne(5918), 11));
            list.Add(str.Substring(IndexMinusOne(5929), 11));
            list.Add(str.Substring(IndexMinusOne(5940), 11));
            list.Add(str.Substring(IndexMinusOne(5951), 8));
            list.Add(str.Substring(IndexMinusOne(5959), 1));
            list.Add(str.Substring(IndexMinusOne(5960), 8));
            list.Add(str.Substring(IndexMinusOne(5968), 10));
            list.Add(str.Substring(IndexMinusOne(5978), 11));
            list.Add(str.Substring(IndexMinusOne(5989), 3));
            list.Add(str.Substring(IndexMinusOne(5992), 2));
            list.Add(str.Substring(IndexMinusOne(5994), 30));
            list.Add(str.Substring(IndexMinusOne(6024), 16));
            list.Add(str.Substring(IndexMinusOne(6040), 60));
            list.Add(str.Substring(IndexMinusOne(6100), 60));
            list.Add(str.Substring(IndexMinusOne(6160), 60));
            list.Add(str.Substring(IndexMinusOne(6220), 280));
            list.Add(str.Substring(IndexMinusOne(6500), 1));
            list.Add(str.Substring(IndexMinusOne(6501), 11));
            list.Add(str.Substring(IndexMinusOne(6512), 8));
            list.Add(str.Substring(IndexMinusOne(6520), 1));
            list.Add(str.Substring(IndexMinusOne(6521), 70));
            list.Add(str.Substring(IndexMinusOne(6591), 9));
            list.Add(str.Substring(IndexMinusOne(6600), 9));
            list.Add(str.Substring(IndexMinusOne(6609), 9));
            list.Add(str.Substring(IndexMinusOne(6618), 9));
            list.Add(str.Substring(IndexMinusOne(6627), 10));
            list.Add(str.Substring(IndexMinusOne(6637), 8));
            list.Add(str.Substring(IndexMinusOne(6645), 8));
            list.Add(str.Substring(IndexMinusOne(6653), 1));
            list.Add(str.Substring(IndexMinusOne(6654), 11));
            list.Add(str.Substring(IndexMinusOne(6665), 11));
            list.Add(str.Substring(IndexMinusOne(6676), 11));
            list.Add(str.Substring(IndexMinusOne(6687), 11));
            list.Add(str.Substring(IndexMinusOne(6698), 1));
            list.Add(str.Substring(IndexMinusOne(6699), 1));
            list.Add(str.Substring(IndexMinusOne(6700), 2));
            list.Add(str.Substring(IndexMinusOne(6702), 1));
            list.Add(str.Substring(IndexMinusOne(6703), 1));
            list.Add(str.Substring(IndexMinusOne(6704), 4));
            list.Add(str.Substring(IndexMinusOne(6708), 40));
            list.Add(str.Substring(IndexMinusOne(6748), 3));
            list.Add(str.Substring(IndexMinusOne(6751), 40));
            list.Add(str.Substring(IndexMinusOne(6791), 1));
            list.Add(str.Substring(IndexMinusOne(6792), 30));
            list.Add(str.Substring(IndexMinusOne(6822), 28));
            list.Add(str.Substring(IndexMinusOne(6850), 28));
            list.Add(str.Substring(IndexMinusOne(6878), 30));
            list.Add(str.Substring(IndexMinusOne(6908), 8));
            list.Add(str.Substring(IndexMinusOne(6916), 1));
            list.Add(str.Substring(IndexMinusOne(6917), 3));
            list.Add(str.Substring(IndexMinusOne(6920), 10));
            list.Add(str.Substring(IndexMinusOne(6930), 3));
            list.Add(str.Substring(IndexMinusOne(6933), 8));
            list.Add(str.Substring(IndexMinusOne(6941), 1));
            list.Add(str.Substring(IndexMinusOne(6942), 1));
            list.Add(str.Substring(IndexMinusOne(6943), 1));
            list.Add(str.Substring(IndexMinusOne(6944), 1));
            list.Add(str.Substring(IndexMinusOne(6945), 1));
            list.Add(str.Substring(IndexMinusOne(6946), 1));
            list.Add(str.Substring(IndexMinusOne(6947), 1));
            list.Add(str.Substring(IndexMinusOne(6948), 1));
            list.Add(str.Substring(IndexMinusOne(6949), 1));
            list.Add(str.Substring(IndexMinusOne(6950), 1));
            list.Add(str.Substring(IndexMinusOne(6951), 1));
            list.Add(str.Substring(IndexMinusOne(6952), 1));
            list.Add(str.Substring(IndexMinusOne(6953), 1));
            list.Add(str.Substring(IndexMinusOne(6954), 1));
            list.Add(str.Substring(IndexMinusOne(6955), 1));
            list.Add(str.Substring(IndexMinusOne(6956), 280));
            list.Add(str.Substring(IndexMinusOne(7236), 1));
            list.Add(str.Substring(IndexMinusOne(7237), 8));
            list.Add(str.Substring(IndexMinusOne(7245), 4));
            list.Add(str.Substring(IndexMinusOne(7249), 1));
            list.Add(str.Substring(IndexMinusOne(7250), 8));
            list.Add(str.Substring(IndexMinusOne(7258), 4));
            list.Add(str.Substring(IndexMinusOne(7262), 1));
            list.Add(str.Substring(IndexMinusOne(7263), 8));
            list.Add(str.Substring(IndexMinusOne(7271), 4));
            list.Add(str.Substring(IndexMinusOne(7275), 1));
            list.Add(str.Substring(IndexMinusOne(7276), 8));
            list.Add(str.Substring(IndexMinusOne(7284), 4));
            list.Add(str.Substring(IndexMinusOne(7288), 1));
            list.Add(str.Substring(IndexMinusOne(7289), 4));
            list.Add(str.Substring(IndexMinusOne(7293), 1));
            list.Add(str.Substring(IndexMinusOne(7294), 1));
            list.Add(str.Substring(IndexMinusOne(7295), 2));
            list.Add(str.Substring(IndexMinusOne(7297), 1));
            list.Add(str.Substring(IndexMinusOne(7298), 4));
            list.Add(str.Substring(IndexMinusOne(7302), 40));
            list.Add(str.Substring(IndexMinusOne(7342), 3));
            list.Add(str.Substring(IndexMinusOne(7345), 40));
            list.Add(str.Substring(IndexMinusOne(7385), 1));
            list.Add(str.Substring(IndexMinusOne(7386), 1));
            list.Add(str.Substring(IndexMinusOne(7387), 30));
            list.Add(str.Substring(IndexMinusOne(7417), 28));
            list.Add(str.Substring(IndexMinusOne(7445), 28));
            list.Add(str.Substring(IndexMinusOne(7473), 1));
            list.Add(str.Substring(IndexMinusOne(7474), 1));
            list.Add(str.Substring(IndexMinusOne(7475), 2));
            list.Add(str.Substring(IndexMinusOne(7477), 1));
            list.Add(str.Substring(IndexMinusOne(7478), 4));
            list.Add(str.Substring(IndexMinusOne(7482), 40));
            list.Add(str.Substring(IndexMinusOne(7522), 3));
            list.Add(str.Substring(IndexMinusOne(7525), 40));
            list.Add(str.Substring(IndexMinusOne(7565), 1));
            list.Add(str.Substring(IndexMinusOne(7566), 1));
            list.Add(str.Substring(IndexMinusOne(7567), 30));
            list.Add(str.Substring(IndexMinusOne(7597), 28));
            list.Add(str.Substring(IndexMinusOne(7625), 28));
            list.Add(str.Substring(IndexMinusOne(7653), 1));
            list.Add(str.Substring(IndexMinusOne(7654), 1));
            list.Add(str.Substring(IndexMinusOne(7655), 2));
            list.Add(str.Substring(IndexMinusOne(7657), 1));
            list.Add(str.Substring(IndexMinusOne(7658), 4));
            list.Add(str.Substring(IndexMinusOne(7662), 40));
            list.Add(str.Substring(IndexMinusOne(7702), 3));
            list.Add(str.Substring(IndexMinusOne(7705), 40));
            list.Add(str.Substring(IndexMinusOne(7745), 1));
            list.Add(str.Substring(IndexMinusOne(7746), 1));
            list.Add(str.Substring(IndexMinusOne(7747), 30));
            list.Add(str.Substring(IndexMinusOne(7777), 28));
            list.Add(str.Substring(IndexMinusOne(7805), 28));
            list.Add(str.Substring(IndexMinusOne(7833), 3));
            list.Add(str.Substring(IndexMinusOne(7836), 2));
            list.Add(str.Substring(IndexMinusOne(7838), 2));
            list.Add(str.Substring(IndexMinusOne(7840), 7));
            list.Add(str.Substring(IndexMinusOne(7847), 1));
            list.Add(str.Substring(IndexMinusOne(7848), 10));
            list.Add(str.Substring(IndexMinusOne(7858), 10));
            list.Add(str.Substring(IndexMinusOne(7868), 8));
            list.Add(str.Substring(IndexMinusOne(7876), 11));
            list.Add(str.Substring(IndexMinusOne(7887), 1));
            list.Add(str.Substring(IndexMinusOne(7888), 3));
            list.Add(str.Substring(IndexMinusOne(7891), 7));
            list.Add(str.Substring(IndexMinusOne(7898), 3));
            list.Add(str.Substring(IndexMinusOne(7901), 1));
            list.Add(str.Substring(IndexMinusOne(7902), 11));
            list.Add(str.Substring(IndexMinusOne(7913), 1));
            list.Add(str.Substring(IndexMinusOne(7914), 8));
            list.Add(str.Substring(IndexMinusOne(7922), 8));
            list.Add(str.Substring(IndexMinusOne(7930), 2));
            list.Add(str.Substring(IndexMinusOne(7932), 1));
            list.Add(str.Substring(IndexMinusOne(7933), 1));
            list.Add(str.Substring(IndexMinusOne(7934), 11));
            list.Add(str.Substring(IndexMinusOne(7945), 1));
            list.Add(str.Substring(IndexMinusOne(7946), 32));
            list.Add(str.Substring(IndexMinusOne(7978), 1));
            list.Add(str.Substring(IndexMinusOne(7979), 1));
            list.Add(str.Substring(IndexMinusOne(7980), 1));
            list.Add(str.Substring(IndexMinusOne(7981), 1));
            list.Add(str.Substring(IndexMinusOne(7982), 19));
            list.Add(str.Substring(IndexMinusOne(8001), 5));
            list.Add(str.Substring(IndexMinusOne(8006), 10));
            list.Add(str.Substring(IndexMinusOne(8016), 5));
            list.Add(str.Substring(IndexMinusOne(8021), 10));
            list.Add(str.Substring(IndexMinusOne(8031), 5));
            list.Add(str.Substring(IndexMinusOne(8036), 10));
            list.Add(str.Substring(IndexMinusOne(8046), 15));
            list.Add(str.Substring(IndexMinusOne(8061), 4));
            list.Add(str.Substring(IndexMinusOne(8065), 5));
            list.Add(str.Substring(IndexMinusOne(8070), 2));
            list.Add(str.Substring(IndexMinusOne(8072), 5));
            list.Add(str.Substring(IndexMinusOne(8077), 3));
            list.Add(str.Substring(IndexMinusOne(8080), 5));
            list.Add(str.Substring(IndexMinusOne(8085), 21));
            list.Add(str.Substring(IndexMinusOne(8106), 3));
            list.Add(str.Substring(IndexMinusOne(8109), 1));
            list.Add(str.Substring(IndexMinusOne(8110), 1));
            list.Add(str.Substring(IndexMinusOne(8111), 1));
            list.Add(str.Substring(IndexMinusOne(8112), 1));
            list.Add(str.Substring(IndexMinusOne(8113), 1));
            list.Add(str.Substring(IndexMinusOne(8114), 2));
            list.Add(str.Substring(IndexMinusOne(8116), 2));
            list.Add(str.Substring(IndexMinusOne(8118), 2));
            list.Add(str.Substring(IndexMinusOne(8120), 8));
            list.Add(str.Substring(IndexMinusOne(8128), 2));
            list.Add(str.Substring(IndexMinusOne(8130), 2));
            list.Add(str.Substring(IndexMinusOne(8132), 11));
            list.Add(str.Substring(IndexMinusOne(8143), 11));
            list.Add(str.Substring(IndexMinusOne(8154), 11));
            list.Add(str.Substring(IndexMinusOne(8165), 11));
            list.Add(str.Substring(IndexMinusOne(8176), 1));
            list.Add(str.Substring(IndexMinusOne(8177), 1));
            list.Add(str.Substring(IndexMinusOne(8178), 10));
            list.Add(str.Substring(IndexMinusOne(8188), 1));
            list.Add(str.Substring(IndexMinusOne(8189), 1));
            list.Add(str.Substring(IndexMinusOne(8190), 10));
            list.Add(str.Substring(IndexMinusOne(8200), 3));
            list.Add(str.Substring(IndexMinusOne(8203), 1));
            list.Add(str.Substring(IndexMinusOne(8204), 1));
            list.Add(str.Substring(IndexMinusOne(8205), 1));
            list.Add(str.Substring(IndexMinusOne(8206), 8));
            list.Add(str.Substring(IndexMinusOne(8214), 1));
            list.Add(str.Substring(IndexMinusOne(8215), 1));
            list.Add(str.Substring(IndexMinusOne(8216), 1));
            list.Add(str.Substring(IndexMinusOne(8217), 1));
            list.Add(str.Substring(IndexMinusOne(8218), 1));
            list.Add(str.Substring(IndexMinusOne(8219), 1));
            list.Add(str.Substring(IndexMinusOne(8220), 1));
            list.Add(str.Substring(IndexMinusOne(8221), 1));
            list.Add(str.Substring(IndexMinusOne(8222), 1));
            list.Add(str.Substring(IndexMinusOne(8223), 1));
            list.Add(str.Substring(IndexMinusOne(8224), 1));
            list.Add(str.Substring(IndexMinusOne(8225), 76));
            list.Add(str.Substring(IndexMinusOne(8301), 1));
            list.Add(str.Substring(IndexMinusOne(8302), 1));
            list.Add(str.Substring(IndexMinusOne(8303), 1));
            list.Add(str.Substring(IndexMinusOne(8304), 1));
            list.Add(str.Substring(IndexMinusOne(8305), 30));
            list.Add(str.Substring(IndexMinusOne(8335), 60));
            list.Add(str.Substring(IndexMinusOne(8395), 30));
            list.Add(str.Substring(IndexMinusOne(8425), 60));
            list.Add(str.Substring(IndexMinusOne(8485), 30));
            list.Add(str.Substring(IndexMinusOne(8515), 60));
            list.Add(str.Substring(IndexMinusOne(8575), 1));
            list.Add(str.Substring(IndexMinusOne(8576), 1));
            list.Add(str.Substring(IndexMinusOne(8577), 30));
            list.Add(str.Substring(IndexMinusOne(8607), 60));
            list.Add(str.Substring(IndexMinusOne(8667), 1));
            list.Add(str.Substring(IndexMinusOne(8668), 30));
            list.Add(str.Substring(IndexMinusOne(8698), 60));
            list.Add(str.Substring(IndexMinusOne(8758), 1));
            list.Add(str.Substring(IndexMinusOne(8759), 30));
            list.Add(str.Substring(IndexMinusOne(8789), 60));
            list.Add(str.Substring(IndexMinusOne(8849), 1));
            list.Add(str.Substring(IndexMinusOne(8850), 1));
            list.Add(str.Substring(IndexMinusOne(8851), 2));
            list.Add(str.Substring(IndexMinusOne(8853), 6));
            list.Add(str.Substring(IndexMinusOne(8859), 1));
            list.Add(str.Substring(IndexMinusOne(8860), 1));
            list.Add(str.Substring(IndexMinusOne(8861), 1));
            list.Add(str.Substring(IndexMinusOne(8862), 6));
            list.Add(str.Substring(IndexMinusOne(8868), 6));
            list.Add(str.Substring(IndexMinusOne(8874)));


            return list;
        }
        #endregion

        #region 10000byteファイルのタイトル
        public static List<string> TableTitle()
        {
            List<string> list = new List<string>();
            list.Add("1-証券番号-");
            list.Add("2-入力ｼｽﾃﾑ-");
            list.Add("3-書類送付区分-");
            list.Add("4-入金確定区分-");
            list.Add("5-CWA情報-");
            list.Add("6-〃-着金日");
            list.Add("7-〃-SUSPENSENO");
            list.Add("8-〃-確定入金額");
            list.Add("9-〃-確定前納回数");
            list.Add("10-〃-確定前納保険料");
            list.Add("11-更新回数情報-申込書更新回数");
            list.Add("12-〃-告知書更新回数");
            list.Add("13-〃-領収証更新回数");
            list.Add("14-〃-口振書更新回数");
            list.Add("15-〃-報告書更新回数");
            list.Add("16-日付情報-申込日");
            list.Add("17-〃-作成日");
            list.Add("18-〃-契約日(予定)");
            list.Add("19-〃-責任開始日");
            list.Add("20-〃-募集月");
            list.Add("21-契約者-性別");
            list.Add("22-〃-ｶﾅ氏名");
            list.Add("23-〃-漢字氏名");
            list.Add("24-〃-生年月日");
            list.Add("25-〃-続柄");
            list.Add("26-〃-電話番号");
            list.Add("27-〃-郵便番号");
            list.Add("28-〃-ｶﾅ住所1");
            list.Add("29-〃-ｶﾅ住所2");
            list.Add("30-〃-ｶﾅ住所3");
            list.Add("31-〃-E-Mail");
            list.Add("32-〃-職業ｺｰﾄﾞ");
            list.Add("33-契約者報告書-年収");
            list.Add("34-〃-他社分契約");
            list.Add("35-〃-給付歴");
            list.Add("36-〃-保険金額(他)");
            list.Add("37-〃-給付金日額(他)");
            list.Add("38-UL用-保険期間(計算ｴﾝｼﾞﾝ)");
            list.Add("39-〃-払込期間(計算ｴﾝｼﾞﾝ)");
            list.Add("40-〃-身長");
            list.Add("41-〃-体重");
            list.Add("42-〃-体重増減");
            list.Add("43-〃-体重増減");
            list.Add("44-〃-体重増減値");
            list.Add("45-〃-危険な趣味");
            list.Add("46-〃-不合格歴,条件付,取消歴");
            list.Add("47-FA用円入金額-円入金額");
            list.Add("48-FA用年金受取人-年金受取人指定ｺｰﾄﾞ");
            list.Add("49-FA用報告書(1)-申込経路");
            list.Add("50-〃-被保険者面接有無");
            list.Add("51-〃-被保険者面接日");
            list.Add("52-〃-被保険者面接場所");
            list.Add("53-報告書等-銀行融資先契約･法人代理店自己契約");
            list.Add("54-FA用報告書(1)-契約者の年収");
            list.Add("55-〃-契約者の金融資産");
            list.Add("56-FILLER-");
            list.Add("57-被保険者-年齢欄の省略有無(FSC)");
            list.Add("58-〃-性別");
            list.Add("59-〃-ｶﾅ氏名");
            list.Add("60-〃-漢字氏名");
            list.Add("61-〃-生年月日");
            list.Add("62-〃-年齢");
            list.Add("63-〃-職業ｺｰﾄﾞ");
            list.Add("64-保険金受取人1-受取人種類");
            list.Add("65-〃-性別");
            list.Add("66-〃-ｶﾅ氏名");
            list.Add("67-〃-漢字氏名");
            list.Add("68-〃-続柄");
            list.Add("69-〃-割合");
            list.Add("70-保険金受取人2～保険金受取人12-");
            list.Add("71-UL用-CV金額");
            list.Add("72-〃-任意P");
            list.Add("73-保険種類-S建,P建");
            list.Add("74-主契約-BASE CODE ");
            list.Add("75-主契約-SUB CODE");
            list.Add("76-主契約-据置期間");
            list.Add("77-主契約-据置期間単位");
            list.Add("71-主契約-払込期間");
            list.Add("72-主契約-払込期間単位");
            list.Add("73-主契約-最低保証年金額(整数部)または保険金額(整数部)");
            list.Add("74-主契約-最低保証年金額(小数部)または保険金額(小数部)");
            list.Add("75-主契約-最低保証年金額の単位");
            list.Add("76-特約１-BASE CODE ");
            list.Add("77-特約１-SUB CODE");
            list.Add("78-特約１-据置期間");
            list.Add("79-特約１-据置期間単位");
            list.Add("73-特約１-払込期間");
            list.Add("74-特約１-払込期間単位");
            list.Add("75-特約１-最低保証年金額(整数部)または保険金額(整数部)");
            list.Add("76-特約１-最低保証年金額(小数部)または保険金額(小数部)");
            list.Add("77-特約１-最低保証年金額の単位");
            list.Add("78-特約２-BASE CODE ");
            list.Add("79-特約２-SUB CODE");
            list.Add("80-特約２-据置期間");
            list.Add("81-特約２-据置期間単位");
            list.Add("75-特約２-払込期間");
            list.Add("76-特約２-払込期間単位");
            list.Add("77-特約２-最低保証年金額(整数部)または保険金額(整数部)");
            list.Add("78-特約２-最低保証年金額(小数部)または保険金額(小数部)");
            list.Add("79-特約２-最低保証年金額の単位");
            list.Add("76-特約３以降-口数,保険金額など");
            list.Add("75-主契約-変額保険種類");
            list.Add("76-〃-死亡逓増期間");
            list.Add("77-主契約・特約-確定保障期間／繰延期間");
            list.Add("78-UL用-UL主契約指定P");
            list.Add("79-〃-UL特約P");
            list.Add("80-〃-平準P");
            list.Add("81-MULTI通貨-米ﾄﾞﾙ");
            list.Add("82-〃-ﾕｰﾛ");
            list.Add("83-〃-豪ﾄﾞﾙ");
            list.Add("84-〃-円");
            list.Add("85-SVR保険料-");
            list.Add("86-特別条件特約-保険金削減年数");
            list.Add("87-〃-割増等級");
            list.Add("88-〃-職業による保険料");
            list.Add("89-合計保険料-");
            list.Add("90-後期合計保険料-");
            list.Add("91-配偶者情報-配偶者有無");
            list.Add("92-〃-ｶﾅ氏名");
            list.Add("93-〃-漢字氏名");
            list.Add("94-〃-生年月日");
            list.Add("95-〃-性別");
            list.Add("96-請求入金関係-払い方");
            list.Add("97-〃-払込経路");
            list.Add("98-〃-提出帳票");
            list.Add("99-〃-請求番号");
            list.Add("100-〃-既契約証券番号");
            list.Add("101-(口座情報)-金融機関ｺｰﾄﾞ");
            list.Add("102-〃-金融機関名称ｶﾅ");
            list.Add("103-〃-金融機関名称漢字");
            list.Add("104-〃-支店ｺｰﾄﾞ");
            list.Add("105-〃-支店名称ｶﾅ");
            list.Add("106-〃-支店名称漢字");
            list.Add("107-〃-口座種目");
            list.Add("108-〃-口座番号");
            list.Add("109-〃-口座名義人ｶﾅ");
            list.Add("110-〃-口座名義人漢字");
            list.Add("111-払込経路の補助-AIU領収書番号");
            list.Add("112-〃-AIU契約ﾀｲﾌﾟ");
            list.Add("113-〃-AIU初回領収月数");
            list.Add("114-〃-AIU領収保険料");
            list.Add("115-〃-税理士登録番号");
            list.Add("116-〃-所在地ｺｰﾄﾞ");
            list.Add("117-〃-単協ｺｰﾄﾞ");
            list.Add("118-〃-医師免許区分");
            list.Add("119-〃-医師免許番号");
            list.Add("120-〃-府県ｺｰﾄﾞ");
            list.Add("121-〃-名前");
            list.Add("122-〃-電話番号");
            list.Add("123-〃-郵便番号");
            list.Add("124-〃-ｶﾅ住所1");
            list.Add("125-〃-ｶﾅ住所2");
            list.Add("126-〃-ｶﾅ住所3");
            list.Add("127-入金情報-入金区分");
            list.Add("128-〃-領収日");
            list.Add("129-〃-領収証事由");
            list.Add("130-〃-領収額");
            list.Add("131-〃-領収証番号");
            list.Add("132-〃-領収証使用件数");
            list.Add("133-〃-前納回数");
            list.Add("134-〃-前納保険料");
            list.Add("135-〃-合計領収額");
            list.Add("136-〃-ｶｰﾄﾞ会社");
            list.Add("137-〃-会員番号");
            list.Add("138-〃-有効期限");
            list.Add("139-〃-会員氏名");
            list.Add("140-〃-承認番号");
            list.Add("141-〃-利用表番号");
            list.Add("142-〃-利用額");
            list.Add("143-〃-合計利用額");
            list.Add("144-〃-利用日(ｵｰｿﾘ日)");
            list.Add("145-保全ｺｰﾄﾞ情報-保全ｺｰﾄﾞ");
            list.Add("146-〃-原契約証券番号");
            list.Add("147-〃-構成員除外者氏名");
            list.Add("148-〃-部位不担保区分1");
            list.Add("149-〃-部位不担保機関指定1");
            list.Add("150-〃-部位不担保機関1");
            list.Add("151-〃-部位不担保種類1");
            list.Add("152-〃-部位不担保ｺｰﾄﾞ1");
            list.Add("153-部位不担保2～部位不担保5-");
            list.Add("154-払込経路の補助-ｼﾃｨﾊﾞﾝｸ顧客番号");
            list.Add("155-本社内判定-成立確認");
            list.Add("156-〃-証券発送停止");
            list.Add("157-〃-証券印刷項目停止");
            list.Add("158-〃-領収証確認ﾌﾗｸﾞ");
            list.Add("159-払込経路の補助-既契約同一口座申込書印刷");
            list.Add("160-WithCash振込区分-");
            list.Add("161-口座名義人-契約者との続柄");
            list.Add("162-FA用第二回入金情報-第二回円入金額");
            list.Add("163-〃-第二回領収日");
            list.Add("164-FILLER-");
            list.Add("165-FA用報告書(2)-契約者確認方法");
            list.Add("166-ライフプランニング-販売コード");
            list.Add("167-告知書-出生時身長");
            list.Add("168-〃-出生時体重");
            list.Add("169-〃-胎在週数");
            list.Add("170-その他-取扱ﾀｲﾌﾟ");
            list.Add("171-〃-証券ﾀｲﾌﾟ");
            list.Add("172-〃-団体ｺｰﾄﾞ");
            list.Add("173-〃-社員ｺｰﾄﾞ");
            list.Add("174-〃-所属ｺｰﾄﾞ");
            list.Add("175-〃-払込に関する特約");
            list.Add("176-〃-証券送付先");
            list.Add("177-〃-告知書の有無");
            list.Add("178-〃-告知日");
            list.Add("179-〃-告知事項はい回答");
            list.Add("180-〃-診査区分");
            list.Add("181-〃-医師区分");
            list.Add("182-〃-嘱託医ｺｰﾄﾞ");
            list.Add("183-〃-面接士区分");
            list.Add("184-〃-面接士ｺｰﾄﾞ");
            list.Add("185-〃-契約者現住所");
            list.Add("186-〃-被保険者現住所");
            list.Add("187-〃-22歳未満の子供の人数");
            list.Add("188-〃-喫煙情報");
            list.Add("189-〃-喫煙本数");
            list.Add("190-〃-営業ｺｰﾄﾞ");
            list.Add("191-〃-質権有無（Restrict　Code）");
            list.Add("192-〃-同時契約件数");
            list.Add("193-〃-事前査定");
            list.Add("194-〃-募集者ｺｰﾄﾞ");
            list.Add("195-〃-共同募集者ｺｰﾄﾞ");
            list.Add("196-〃-AFCｺｰﾄﾞ");
            list.Add("197-FA用報告書(3)-契約者の投資の経験");
            list.Add("198-FILLER-");
            list.Add("199-報告書等-身長");
            list.Add("200-〃-体重");
            list.Add("201-〃-体重増減");
            list.Add("202-〃-申込動機(保険の目的)");
            list.Add("203-〃-申込経路");
            list.Add("204-〃-被保険者面接日");
            list.Add("205-〃-被保険者面接場所");
            list.Add("206-〃-申込書不備");
            list.Add("207-〃-危険な趣味");
            list.Add("208-〃-年収");
            list.Add("209-〃-上記以外の職務");
            list.Add("210-〃-副業 ");
            list.Add("211-〃-不合格歴,条件付,取消歴");
            list.Add("212-〃-他社分契約");
            list.Add("213-〃-給付歴");
            list.Add("214-〃-保険金額(他)");
            list.Add("215-〃-給付金日額(他)");
            list.Add("216-〃-飲酒の有無");
            list.Add("217-〃-飲酒種類");
            list.Add("218-〃-飲酒量");
            list.Add("219-〃-被保険者が家計中心者");
            list.Add("220-〃-家計中心者続柄");
            list.Add("221-〃-家計中心者年収");
            list.Add("222-〃-家計中心者付保額");
            list.Add("223-〃-家族歴(2人以上60歳未満病死)");
            list.Add("224-〃-契約者面接日");
            list.Add("225-〃-契約者面接場所");
            list.Add("226-〃-事業保険設立年月");
            list.Add("227-〃-事業保険資本金");
            list.Add("228-〃-事業保険従業員人数");
            list.Add("229-〃-事業保険年商");
            list.Add("230-〃-事業保険税引利益");
            list.Add("231-〃-事業保険社内規定");
            list.Add("232-〃-事業保険新規社内規定");
            list.Add("233-〃-事業保険下位者の付保");
            list.Add("234-〃-被保険者確認方法");
            list.Add("235-〃-飲酒量");
            list.Add("236-〃-飲酒量");
            list.Add("237-〃-飲酒量");
            list.Add("238-〃-申込動機(保険の目的)2");
            list.Add("239-〃-事業保険被保険者人数");
            list.Add("240-〃-被保険者体重増減");
            list.Add("241-〃-被保険者体重増減値");
            list.Add("242-〃-上記以外の職務詳細");
            list.Add("243-〃-副業詳細");
            list.Add("244-〃-実父死因");
            list.Add("245-〃-実父死亡年齢");
            list.Add("246-〃-実母死因");
            list.Add("247-〃-実母死亡年齢");
            list.Add("248-〃-兄弟姉妹1死因");
            list.Add("249-〃-兄弟姉妹1死亡年齢");
            list.Add("250-〃-兄弟姉妹2死因");
            list.Add("251-〃-兄弟姉妹2死亡年齢");
            list.Add("252-〃-兄弟姉妹3死因");
            list.Add("253-〃-兄弟姉妹3死亡年齢");
            list.Add("254-診査区分-看護担当者ｺｰﾄﾞ");
            list.Add("255-報告書等-事業保険災害死亡時S");
            list.Add("256-〃-勤続2年以上ﾀｸｼｰ以外");
            list.Add("257-〃-災害死亡保険金額(他)");
            list.Add("258-〃-副業の職業ｺｰﾄﾞ");
            list.Add("259-〃-以外の職務職業ｺｰﾄﾞ");
            list.Add("260-告知書-被保険者性別");
            list.Add("261-報告書等-米国人示唆情報有");
            list.Add("262-FA用報告書(4)-強制Pending査定結果");
            list.Add("263-保全コード情報-追加分");
            list.Add("264-ソフト版ALEX-");
            list.Add("265-本社内処理-ｻﾏﾘｰTR強制'修'書出し");
            list.Add("266-〃-ｻﾏﾘｰTR強制'CWA'変更");
            list.Add("267-FA用報告書(5)-契約者保険の目的");
            list.Add("268-FSC-確認日有無");
            list.Add("269--構成員契約の該当");
            list.Add("270-子供情報１-ｶﾅ氏名");
            list.Add("271-〃-漢字氏名");
            list.Add("272-〃-生年月日");
            list.Add("273-〃-性別");
            list.Add("274-子供情報２-ｶﾅ氏名");
            list.Add("275-〃-漢字氏名");
            list.Add("276-〃-生年月日");
            list.Add("277-〃-性別");
            list.Add("278-子供情報３-ｶﾅ氏名");
            list.Add("279-〃-漢字氏名");
            list.Add("280-〃-生年月日");
            list.Add("281-〃-性別");
            list.Add("282-子供情報４-ｶﾅ氏名");
            list.Add("283-〃-漢字氏名");
            list.Add("284-〃-生年月日");
            list.Add("285-〃-性別");
            list.Add("286-契約者通信先-漢字住所1");
            list.Add("287-〃-漢字住所2");
            list.Add("288-〃-漢字住所3");
            list.Add("289-Debit情報-ご利用日");
            list.Add("290-〃-処理通番");
            list.Add("291-〃-処理件数");
            list.Add("292-〃-利用額");
            list.Add("293-〃-合計利用額");
            list.Add("294-親1情報-親1有無");
            list.Add("295-〃-ｶﾅ氏名");
            list.Add("296-〃-漢字氏名");
            list.Add("297-〃-生年月日");
            list.Add("298-〃-性別");
            list.Add("299-親2情報-親2有無");
            list.Add("300-〃-ｶﾅ氏名");
            list.Add("301-〃-漢字氏名");
            list.Add("302-〃-生年月日");
            list.Add("303-〃-性別");
            list.Add("304-FA用-NDPｴｰｼﾞｪﾝﾄｺｰﾄﾞ1");
            list.Add("305-〃-NDPｴｰｼﾞｪﾝﾄｺｰﾄﾞ2");
            list.Add("306-〃-法人代理店使用ｺｰﾄﾞ");
            list.Add("307-〃-積立利率･整数部");
            list.Add("308-〃-積立利率･少数部");
            list.Add("309-FSC-BANK用-BANK情報の有無");
            list.Add("310-〃-大和証券-出金経路");
            list.Add("311-申込人-郵便番号");
            list.Add("312-FSC-BANK用-入金可否");
            list.Add("313-FSC-BANK用-大和証券-FAX受付日");
            list.Add("314-〃-現物書類有無(大和･City)");
            list.Add("315-FILLER-");
            list.Add("316-申込人-性別");
            list.Add("317-主契約,特約 (特約の20番目～)-口数,保険金額など");
            list.Add("318-FSC-BANK用-保険種類ｺｰﾄﾞ");
            list.Add("319-〃-CIFｺｰﾄﾞ");
            list.Add("320-SW申込書-SW申込書の有無");
            list.Add("321-〃-SW指定契約合計保険料合算");
            list.Add("322-〃-定期取崩額");
            list.Add("323-〃-請求停止ﾌﾗｸﾞ");
            list.Add("324-〃-SW指定契約明細");
            list.Add("325-FILLER-");
            list.Add("326-申込人-漢字氏名");
            list.Add("327-〃-続柄");
            list.Add("328-〃-生年月日");
            list.Add("329-〃-ｶﾅ住所1");
            list.Add("330-〃-ｶﾅ住所2");
            list.Add("331-〃-ｶﾅ住所3");
            list.Add("332-FSC Banc拡大-契約者勤務先");
            list.Add("333-〃-親権者有無(契約者)");
            list.Add("334-〃-親権者該当区分(契約者)");
            list.Add("335-〃-親権者有無(被保険者)");
            list.Add("336-〃-親権者該当区分(被保険者)");
            list.Add("337-〃-確認日");
            list.Add("338-〃-報告日");
            list.Add("339-FSC-BANK用-主-金融機関ｺｰﾄﾞ");
            list.Add("340-〃-主-支店/分室ｺｰﾄﾞ");
            list.Add("341-〃-主-行員ｺｰﾄﾞ");
            list.Add("342-〃-主-募集人登録番号");
            list.Add("343-〃-主-募集人正規L1ｺｰﾄﾞ");
            list.Add("344-〃-従-金融機関ｺｰﾄﾞ");
            list.Add("345-〃-従-支店/分室ｺｰﾄﾞ");
            list.Add("346-〃-従-行員ｺｰﾄﾞ");
            list.Add("347-〃-従-募集人登録番号");
            list.Add("348-〃-従-募集人正規L1ｺｰﾄﾞ");
            list.Add("349-〃-従-ﾀﾞﾐｰL1ｺｰﾄﾞ");
            list.Add("350-MCFA用?@-BASE CODE");
            list.Add("351-MCFA用-SUB CODE");
            list.Add("352-MCFA用-据置期間");
            list.Add("353-MCFA用-据置期間単位");
            list.Add("354-MCFA用-払込期間 ");
            list.Add("355-MCFA用-払込期間単位");
            list.Add("356-MCFA用-最低保証年金額(整数部)");
            list.Add("357-MCFA用-最低保証年金額(少数部)");
            list.Add("358-MCFA用-金額の単位");
            list.Add("359-MCFA用-積立利率保証期間      ");
            list.Add("360-MCFA用-合計保険料       ");
            list.Add("361-MCFA用-領収日       ");
            list.Add("362-MCFA用-領収額             ");
            list.Add("363-MCFA用-円入金額 ");
            list.Add("364-MCFA用-第二回円入金額");
            list.Add("365-MCFA用-第二回領収日");
            list.Add("366-MCFA用-CWA情報 ");
            list.Add("367-MCFA用-着金日 ");
            list.Add("368-MCFA用-SUSPENSENO　　");
            list.Add("369-MCFA用-確定入金額   　");
            list.Add("370-MCFA用-積立利率･整数部   　");
            list.Add("371-MCFA用-積立利率･少数部     　");
            list.Add("372-MCFA用?A-BASE CODE");
            list.Add("373-MCFA用-SUB CODE");
            list.Add("374-MCFA用-据置期間");
            list.Add("375-MCFA用-据置期間単位");
            list.Add("376-MCFA用-払込期間 ");
            list.Add("377-MCFA用-払込期間単位");
            list.Add("378-MCFA用-最低保証年金額(整数部)");
            list.Add("379-MCFA用-最低保証年金額(少数部)");
            list.Add("380-MCFA用-金額の単位");
            list.Add("381-MCFA用-積立利率保証期間      ");
            list.Add("382-MCFA用-合計保険料       ");
            list.Add("383-MCFA用-領収日       ");
            list.Add("384-MCFA用-領収額             ");
            list.Add("385-MCFA用-円入金額 ");
            list.Add("386-MCFA用-第二回円入金額");
            list.Add("387-MCFA用-第二回領収日");
            list.Add("388-MCFA用-CWA情報 ");
            list.Add("389-MCFA用-着金日 ");
            list.Add("390-MCFA用-SUSPENSENO　　");
            list.Add("391-MCFA用-確定入金額   　");
            list.Add("392-MCFA用-積立利率･整数部   　");
            list.Add("393-MCFA用-積立利率･少数部     　");
            list.Add("394-MCFA用?B-BASE CODE");
            list.Add("395-MCFA用-SUB CODE");
            list.Add("396-MCFA用-据置期間");
            list.Add("397-MCFA用-据置期間単位");
            list.Add("398-MCFA用-払込期間 ");
            list.Add("399-MCFA用-払込期間単位");
            list.Add("400-MCFA用-最低保証年金額(整数部)");
            list.Add("401-MCFA用-最低保証年金額(少数部)");
            list.Add("402-MCFA用-金額の単位");
            list.Add("403-MCFA用-積立利率保証期間      ");
            list.Add("404-MCFA用-合計保険料       ");
            list.Add("405-MCFA用-領収日       ");
            list.Add("406-MCFA用-領収額             ");
            list.Add("407-MCFA用-円入金額 ");
            list.Add("408-MCFA用-第二回円入金額");
            list.Add("409-MCFA用-第二回領収日");
            list.Add("410-MCFA用-CWA情報 ");
            list.Add("411-MCFA用-着金日 ");
            list.Add("412-MCFA用-SUSPENSENO　　");
            list.Add("413-MCFA用-確定入金額   　");
            list.Add("414-MCFA用-積立利率･整数部   　");
            list.Add("415-MCFA用-積立利率･整数部   　");
            list.Add("416-MCFA用?C-BASE CODE");
            list.Add("417-MCFA用-SUB CODE");
            list.Add("418-MCFA用-据置期間");
            list.Add("419-MCFA用-据置期間単位");
            list.Add("420-MCFA用-払込期間 ");
            list.Add("421-MCFA用-払込期間単位");
            list.Add("422-MCFA用-最低保証年金額(整数部)");
            list.Add("423-MCFA用-最低保証年金額(少数部)");
            list.Add("424-MCFA用-金額の単位");
            list.Add("425-MCFA用-積立利率保証期間      ");
            list.Add("426-MCFA用-合計保険料       ");
            list.Add("427-MCFA用-領収日       ");
            list.Add("428-MCFA用-領収額             ");
            list.Add("429-MCFA用-円入金額 ");
            list.Add("430-MCFA用-第二回円入金額");
            list.Add("431-MCFA用-第二回領収日");
            list.Add("432-MCFA用-CWA情報 ");
            list.Add("433-MCFA用-着金日 ");
            list.Add("434-MCFA用-SUSPENSENO　　");
            list.Add("435-MCFA用-確定入金額   　");
            list.Add("436-MCFA用-積立利率･整数部   　");
            list.Add("437-MCFA用-積立利率･整数部   　");
            list.Add("351-申込人-ｶﾅ氏名");
            list.Add("352-〃-電話番号");
            list.Add("353-〃-漢字住所1");
            list.Add("354-〃-漢字住所2");
            list.Add("355-〃-漢字住所3");
            list.Add("356-〃-ｼｮｰﾄﾈｰﾑ");
            list.Add("357-MCFA用-longevity有無");
            list.Add("358-次期解禁用-SMBC-受付管理ID");
            list.Add("359-〃-報状到着日");
            list.Add("360-〃-被保険者の年収");
            list.Add("361-〃-被保険者勤務先");
            list.Add("362-〃-SVR生存給付金1");
            list.Add("363-〃-SVR生存給付金2");
            list.Add("364-〃-SVR生存給付金3");
            list.Add("365-〃-SVR生存給付金４");
            list.Add("366-ALEX2用-ﾊﾞｰｼﾞｮﾝ");
            list.Add("367-次期解禁用-診査予定日");
            list.Add("368-〃-被保険者生年月日(告知書)");
            list.Add("369-FSC-保険料円入金-保険料円入金特約");
            list.Add("370-〃-保険料円払込額(主契約)");
            list.Add("371-〃-保険料円払込額(ﾕｰﾛ)");
            list.Add("372-〃-保険料円払込額(豪ﾄﾞﾙ)");
            list.Add("373-〃-保険料円払込額(MCFA用)");
            list.Add("374-FSC-定期引出特約-定期引出特約ﾌﾗｸﾞ");
            list.Add("375-〃-定期引出特約の型");
            list.Add("376-〃-分割受取回数");
            list.Add("377-〃-円貨受取");
            list.Add("378-〃-送金口座(通貨種類)");
            list.Add("379-〃-送金口座(金融機関ｺｰﾄﾞ)");
            list.Add("380-〃-送金口座(金融機関名称)");
            list.Add("381-〃-送金口座(支店ｺｰﾄﾞ)");
            list.Add("382-〃-送金口座(支店名称)");
            list.Add("383-〃-送金口座(預金種目)");
            list.Add("384-〃-送金口座(口座番号)");
            list.Add("385-〃-送金口座(口座名義人カナ)");
            list.Add("386-〃-送金口座(口座名義人漢字)");
            list.Add("387-基本情報-第二保全ｺｰﾄﾞ");
            list.Add("388-FSC-約款受領書確認日");
            list.Add("389-MCFA用-円建年金移行特約FLG");
            list.Add("390-〃-円建年金移行特約目標額");
            list.Add("391-FSC-意向確認書-帳票種類No.");
            list.Add("392-〃-ﾊﾞｰｼﾞｮﾝ");
            list.Add("393-〃-確認日");
            list.Add("394-〃-加入目的");
            list.Add("395-〃-収入");
            list.Add("396-〃-金融資産");
            list.Add("397-〃-投資の経験");
            list.Add("398-〃-保険料の原資");
            list.Add("399-〃-項目1");
            list.Add("400-〃-項目2");
            list.Add("401-〃-項目3");
            list.Add("402-〃-項目4");
            list.Add("403-〃-項目5");
            list.Add("404-〃-項目6");
            list.Add("405-〃-項目7");
            list.Add("406-〃-項目8");
            list.Add("407-〃-項目9");
            list.Add("408-外部入力区分-");
            list.Add("409-保険 ID(PLAN ID)-");
            list.Add("410-犯罪収益移転防止確認-契約者＜個人＞ 確認書類");
            list.Add("411-　　〃-契約者＜個人＞ 確認日");
            list.Add("412-　　〃-契約者＜個人＞ 確認時刻");
            list.Add("413-　　〃-契約者＜法人＞ 確認書類");
            list.Add("414-　　〃-契約者＜法人＞ 確認日");
            list.Add("415-　　〃-契約者＜法人＞ 確認時刻");
            list.Add("416-　　〃-取引担当者 確認書類");
            list.Add("417-　　〃-取引担当者 確認日");
            list.Add("418-　　〃-取引担当者 確認時刻");
            list.Add("419-　　〃-親権者等 確認書類");
            list.Add("420-　　〃-親権者等 確認日");
            list.Add("421-　　〃-親権者等 確認時刻");
            list.Add("422-年金支払特約情報-特約付加有無");
            list.Add("423-　　〃-年金種類");
            list.Add("424-FSC-定期引出特約NewMCFA（主契約）-定期引出特約ﾌﾗｸﾞ");
            list.Add("425-FSC-定期引出特約NewMCFA（主契約）-定期引出特約の型");
            list.Add("426-FSC-定期引出特約NewMCFA（主契約）-分割受取回数");
            list.Add("427-FSC-定期引出特約NewMCFA（主契約）-円貨受取");
            list.Add("429-FSC-定期引出特約NewMCFA（主契約）-送金先金融機関ｺｰﾄﾞ");
            list.Add("430-FSC-定期引出特約NewMCFA（主契約）-送金先金融機関名称");
            list.Add("431-FSC-定期引出特約NewMCFA（主契約）-送金先支店ｺｰﾄﾞ");
            list.Add("432-FSC-定期引出特約NewMCFA（主契約）-送金先支店名称");
            list.Add("433-FSC-定期引出特約NewMCFA（主契約）-送金先口座種目");
            list.Add("428-FSC-定期引出特約NewMCFA（主契約）-送金通貨種類");
            list.Add("434-FSC-定期引出特約NewMCFA（主契約）-送金先口座番号");
            list.Add("435-FSC-定期引出特約NewMCFA（主契約）-送金先口座名義人カナ");
            list.Add("436-FSC-定期引出特約NewMCFA（主契約）-送金先口座名義人漢字");
            list.Add("424-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-定期引出特約ﾌﾗｸﾞ");
            list.Add("425-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-定期引出特約の型");
            list.Add("426-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-分割受取回数");
            list.Add("427-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-円貨受取");
            list.Add("428-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先金融機関ｺｰﾄﾞ");
            list.Add("429-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先金融機関名称");
            list.Add("430-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先支店ｺｰﾄﾞ");
            list.Add("431-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先支店名称");
            list.Add("432-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先口座種目");
            list.Add("433-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金通貨種類");
            list.Add("434-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先口座番号");
            list.Add("435-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先口座名義人カナ");
            list.Add("436-FSC-定期引出特約NewMCFA(ﾕｰﾛ)-送金先口座名義人漢字");
            list.Add("437-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-定期引出特約ﾌﾗｸﾞ");
            list.Add("438-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-定期引出特約の型");
            list.Add("439-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-分割受取回数");
            list.Add("440-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-円貨受取");
            list.Add("441-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先金融機関ｺｰﾄﾞ");
            list.Add("442-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先金融機関名称");
            list.Add("443-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先支店ｺｰﾄﾞ");
            list.Add("444-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先支店名称");
            list.Add("445-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先口座種目");
            list.Add("446-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金通貨種類");
            list.Add("447-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先口座番号");
            list.Add("448-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先口座名義人カナ");
            list.Add("449-FSC-定期引出特約NewMCFA(豪ﾄﾞﾙ)-送金先口座名義人漢字");
            list.Add("450-高齢者募集報告書-同席者年齢");
            list.Add("451-〃-同席者続柄");
            list.Add("452-〃-立会者続柄");
            list.Add("453-〃-募集者コード");
            list.Add("454-告知書-検査項目");
            list.Add("455-外貨入金特約-豪ドル・外貨払込額（米ドル）");
            list.Add("456-〃-領収証情報　初回・SUSNO");
            list.Add("457-〃-領収証情報　初回・領収日");
            list.Add("458-〃-領収証情報　初回・入金額");
            list.Add("459-後継ﾌﾟﾗﾝ-元契約分受取金額の充当");
            list.Add("460-FSC-BANK(野村證券)用-野村證券-取引店ｺｰﾄﾞ");
            list.Add("461-　 〃-野村證券-口座番号");
            list.Add("462-　 〃-野村證券-係");
            list.Add("463--外貨入金特約有無");
            list.Add("464--初回米ドル入金額（保険料）");
            list.Add("465-告知書-ガン告知");
            list.Add("466-高齢者募集報告書-同席日");
            list.Add("467-〃-商品等説明日");
            list.Add("468-〃-上席者同席希望の案内");
            list.Add("469-意向確認書-意向確認書有無");
            list.Add("470-〃-意向確認書内容");
            list.Add("471-外貨入金特約-外貨払込額");
            list.Add("472-報告書など-意向の把握状況");
            list.Add("473--申込経路　その他内容");
            list.Add("474--紹介者");
            list.Add("475--紹介報酬の有無");
            list.Add("476--紹介者による商品説明の有無");
            list.Add("477--乗換募集");
            list.Add("469-FILLER-");
            list.Add("470-ﾙｰﾙｴﾝｼﾞﾝ用-査定者ID(最新)");
            list.Add("471-〃-査定結果(最新)");
            list.Add("472-〃-査定者ID(ひとつ前)");
            list.Add("473-〃-査定結果(ひとつ前)");
            list.Add("474-〃-査定者ID(ふたつ前)");
            list.Add("475-〃-査定結果(ふたつ前)");
            list.Add("476-〃-診査情報(最新)");
            list.Add("477-〃-延期期間(最新)");
            list.Add("478-〃-疾病ｺｰﾄﾞ(最新)");
            list.Add("479-〃-業務ｺｰﾄﾞ");
            list.Add("480-〃-代理店ｺｰﾄﾞ");
            list.Add("481-〃-商品ｺｰﾄﾞ");
            list.Add("482-〃-法人業務用支店ｺｰﾄﾞ");
            list.Add("483-〃-OCN");
            list.Add("484-〃-販売名称ｺｰﾄﾞ");
            list.Add("485-〃-米ﾄﾞﾙ");
            list.Add("486-〃-ﾕｰﾛ");
            list.Add("487-〃-豪ﾄﾞﾙ");
            list.Add("488-〃-円");
            list.Add("489-〃-送金明細書STATUS");
            list.Add("490-〃-受取人NameNoMax");
            list.Add("491-〃-配偶者NameNo");
            list.Add("492-〃-申込人No");
            list.Add("493-〃-保険 ID");
            list.Add("494-〃-領収証用保険種類ｺｰﾄﾞ");
            list.Add("495-〃-払込経路枝番");
            list.Add("496-〃-基本保険金額S");
            list.Add("497-〃-増加死亡保険金額S");
            list.Add("498-〃-円入金額");
            list.Add("499-〃-第二回円入金額");
            list.Add("500-〃-新変更特約申込書有無");
            list.Add("501-〃-新変更特約送信時OPI");
            list.Add("502-〃-ﾌﾟﾛｽﾍﾟｸﾄ??");
            list.Add("503-〃-配偶者情報反映ﾌﾗｸﾞ");
            list.Add("504-〃-個人／法人区分");
            list.Add("505-FSC-意向確認書-帳票種類No（適合性確認書）");
            list.Add("506--バージョン（適合性確認書）");
            list.Add("507--適合性確認書有無");
            list.Add("508--意向把握は実施済");
            list.Add("509--当初意向把握日フラグ");
            list.Add("510--当初意向の把握日");
            list.Add("511--当初_死亡に備えての保障");
            list.Add("512--当初_病気・ケガに備えての保障");
            list.Add("513--当初_ガンに備えての保障");
            list.Add("514--当初_介護の保障");
            list.Add("515--当初_教育・老後に備えての保障");
            list.Add("516--当初_資産の運用");
            list.Add("517--当初_不備");
            list.Add("518--当初_回答なし");
            list.Add("519--当初_項目なし");
            list.Add("520--当初_貯蓄分野について");
            list.Add("521--意向振り返り");
            list.Add("505-FILLER-");
            list.Add("506-報告書など-契約者居住地国");
            list.Add("507--契約者(個人) 本人確認書類２");
            list.Add("508--取引担当者   本人確認書類２");
            list.Add("509--親権者       本人確認書類２");
            list.Add("510--契約者(法人) 実質的支配者カナ1");
            list.Add("511--契約者(法人) 実質的支配者氏名1");
            list.Add("512--契約者(法人) 実質的支配者カナ2");
            list.Add("513--契約者(法人) 実質的支配者氏名2");
            list.Add("514--契約者(法人) 実質的支配者カナ3");
            list.Add("515--契約者(法人) 実質的支配者氏名3");
            list.Add("516--CRS届出書 届出書有無");
            list.Add("517--CRS届出書 本店の居住地国");
            list.Add("518--CRS届出書 実質的支配者カナ1");
            list.Add("519--CRS届出書 実質的支配者氏名1");
            list.Add("520--CRS届出書 実質的支配者居住地国１");
            list.Add("521--CRS届出書 実質的支配者カナ2");
            list.Add("522--CRS届出書 実質的支配者氏名2");
            list.Add("523--CRS届出書 実質的支配者居住地国２");
            list.Add("524--CRS届出書 実質的支配者カナ3");
            list.Add("525--CRS届出書 実質的支配者氏名3");
            list.Add("526--CRS届出書 実質的支配者居住地国３");
            list.Add("524--「契約者」家計の中心者");
            list.Add("525--家計の中心者と契約者の続柄");
            list.Add("526--契約者の家計の中心者の年収（金額）");
            list.Add("527--契約者の家計の中心者の年収（区分）");
            list.Add("528--被保険者の家計の中心者の年収（区分）");
            list.Add("529--特定保険の年齢職業チェック");
            list.Add("530--世帯の金融資産");
            list.Add("531--世帯の年収");
            list.Add("527-FILLER-");

            return list;
        }
        #endregion

        #region　index-1
        public static int IndexMinusOne(int index)
        {
            return index - 1;
        }
        #endregion
    }


    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DataValueDefaultAttribute : Attribute
    {
        private object value;

        public DataValueDefaultAttribute(object value)
        {
            this.value = value;
        }
        public object Value { get { return value; } }
    }
}



