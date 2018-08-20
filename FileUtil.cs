using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;
using Aspose.Cells;
using QiShi_BAS_Common;
using QiShi_BAS_DAL;
using QiShi_BAS_VO;

namespace QiShi_BAS_MVC
{
    public class FileUtil
    {
        /// <summary>
        /// 将obj转换成json
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ToJson(object obj)
        {
            JavaScriptSerializer ser = new JavaScriptSerializer();
            ser.MaxJsonLength = Int32.MaxValue;
            return ser.Serialize(obj);
        }
        /// <summary>
        /// 导出数据
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="datatable"></param>
        /// <param name="error"></param>
        /// <returns></returns>
        public static bool ExcelFileToDataTable(string filepath, out DataTable datatable, out string error)
        {
            error = "";
            datatable = null;
            var p = System.Web.HttpContext.Current.Server.MapPath("~/") + filepath;
            try
            {
                if (File.Exists(p) == false)
                {
                    error = "文件不存在";
                    datatable = null;
                    return false;
                }
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
                workbook.Open(p);
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];
                datatable = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxRow + 1, worksheet.Cells.MaxColumn + 1);
                return true;
            }
            catch (System.Exception e)
            {
                error = e.Message;
                return false;
            }
        }
         
        /// <summary>
        /// 导出投诉统计
        /// </summary>
        /// <param name="dept"></param>
        /// <param name="deptname"></param>
        /// <param name="st"></param>
        /// <param name="et"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public static bool ComplaintoTable(string dept, string deptname, string st, string et , ref string result)
        {
            bool r = false;
            #region 读取数值
            DataTable dt = new DataTable();
            string where = "";
            if (!string.IsNullOrWhiteSpace(dept)) where = where + " and Order_Info.DeptID=" + dept;
            if (st == et)
            {
                if (!string.IsNullOrWhiteSpace(st)) where = where + " and Order_Complaint.ComplainDate='" + st + "'";

            }
            else
            {
                if (!string.IsNullOrWhiteSpace(st)) where = where + " and Order_Complaint.ComplainDate>='" + st + "'";
                if (!string.IsNullOrWhiteSpace(et)) where = where + " and Order_Complaint.ComplainDate<='" + et + "'";
            }
            Order_ComplaintDAL ocdata = new Order_ComplaintDAL();
            DataTable db = ocdata.TableOrder_Complaint1(where);
      
            int c = 1;
            if (db != null && db.Rows.Count > 0)
            {


                dt.Columns.Add("序号", typeof(string));
                dt.Columns.Add("部门", typeof(string));
                dt.Columns.Add("工号", typeof(string));
                dt.Columns.Add("姓名", typeof(string));
                dt.Columns.Add("退费总额", typeof(string));
                dt.Columns.Add("投诉量", typeof(string));
                dt.Columns.Add("取消量", typeof(string)); 
                dt.Columns.Add("话务员责任量", typeof(string));
          
                double d1 = 0, d2 = 0, d3 = 0, d4 = 0;
                ArrayList tempList = new ArrayList();
                foreach (DataRow m in db.Rows)
                {
                    tempList = new ArrayList();
                    tempList.Add(c);
                    tempList.Add(m["DeptName"]);
                    tempList.Add(m["UserID"]);
                    tempList.Add(m["UserName"]);
                    tempList.Add(m["a1"]); d1 += Convert.ToDouble(m["a1"]);
                    tempList.Add(m["a2"]); d2 += Convert.ToDouble(m["a2"]);
                    tempList.Add(m["a3"]); d3 += Convert.ToDouble(m["a3"]);
                    tempList.Add(m["a4"]); d4 += Convert.ToDouble(m["a4"]); 
                    dt.LoadDataRow(tempList.ToArray(), true);
                    c++;
                }
                tempList = new ArrayList();
                tempList.Add(""); 
                tempList.Add("合计");
                tempList.Add("");
                tempList.Add("");
                tempList.Add(d1);
                tempList.Add(d2);
                tempList.Add(d3);
                tempList.Add(d4);
                dt.LoadDataRow(tempList.ToArray(), true);
            }

            #endregion

            result = "";
            try
            { //创建文件夹
                String savePath = "../upload/file/";
                String dirPath = System.Web.HttpContext.Current.Server.MapPath(savePath);
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }
                DateTime now = DateTime.Now;
                String ymd = now.ToString("yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                dirPath += ymd + "/";
                savePath += ymd + "/";
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }
                String newFileName = deptname + "投诉统计表" + st+"至"+et;
                String filepath = dirPath + newFileName + ".xlsx";


                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();

                try
                {
                    if (dt == null)
                    {
                        result = "DataTableToExcel:datatable 为空";
                        return false;
                    }
                    FileInfo info = new FileInfo(filepath);
                    info.Create().Dispose();
                    //为单元格添加样式    
                    Aspose.Cells.Style style = wb.Styles[wb.Styles.Add()];

                    Worksheet sheet = wb.Worksheets[0]; //工作表
                    Cells cells = sheet.Cells;//单元格 
                    //设置居中
                    style.HorizontalAlignment = Aspose.Cells.TextAlignmentType.Center;
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.TopBorder].Color = Color.Black;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].Color = Color.Black;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].Color = Color.Black;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].Color = Color.Black;
                    style.IsTextWrapped = true;//单元格内容自动换行
                    style.Font.IsBold = true;
                    style.Pattern = BackgroundType.Solid;
                    style.ForegroundColor = ColorTranslator.FromHtml("#B6DDE8");
                    style.Font.Color = Color.Black;

                    cells.Merge(0, 0, 1, 8);//合并单元格  
                    cells[0, 0].PutValue(newFileName);//填写内容  
                    cells[0, 0].SetStyle(style);//给单元格关联样式   
                    cells.Merge(0, 8, 1,1);//合并单元格  
                    cells[0, 8].PutValue("");//填写内容  
                    cells.SetRowHeight(0, 25);//设置行高   

                    int rowIndex = 1;
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        DataColumn col = dt.Columns[i];
                        string columnName = col.Caption ?? col.ColumnName;

                        cells[rowIndex, i].PutValue(columnName);
                        cells[rowIndex, i].SetStyle(style);

                    }
                    cells.SetRowHeight(rowIndex, 25);//设置行高 

                    rowIndex++;
                    style.Font.IsBold = false;
                    foreach (DataRow row in dt.Rows)
                    {

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            style.Font.Color = Color.Black;
                            style.ForegroundColor = Color.White;
                            cells[rowIndex, i].PutValue(row[i].ToString());
                            if (rowIndex == dt.Rows.Count + 1)
                            {
                                style.Font.IsBold = true;
                                style.ForegroundColor = ColorTranslator.FromHtml("#B6DDE8");
                            }
                            cells[rowIndex, i].SetStyle(style);
                        }
                        cells.SetRowHeight(rowIndex, 25);//设置行高 
                        rowIndex++;
                    }

                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        sheet.AutoFitColumn(k, 0, dt.Rows.Count);
                    }
                    for (int col = 0; col < cells.Columns.Count; col++)
                    {
                        if (cells.GetColumnWidthPixel(col) > 300) cells.SetColumnWidthPixel(col, 300);
                    }
                    sheet.FreezePanes("A3", 2, dt.Columns.Count);
                    sheet.AutoFilter.Range = "A2:" + ExcelConvert.ToName(dt.Columns.Count - 1) + "2";

                    wb.Save(filepath);
                    result = savePath + newFileName + ".xlsx";
                    return true;
                }
                catch (Exception e)
                {
                    result = result + " DataTableToExcel: " + e.Message;
                    return false;
                }

            }
            catch (System.Exception e)
            {
                result = "失败：" + e.Message;
            }
            return r;
        }
    }
}