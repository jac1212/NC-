using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NCvoucher
{
    class InteropHelper
    {
        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="excelPath">excel 文件路径</param>
        /// <returns></returns>
        public static System.Data.DataTable ReadExcel(int rowNum, string fileName, string sheetName = null, bool isFirstRowColumn = true)
        {
            //定义要返回的datatable对象
            System.Data.DataTable data = new System.Data.DataTable();
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;//设置调用引用的 Excel文件是否可见
            excel.DisplayAlerts = false;
            Workbook wb = null;
            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                if (!File.Exists(fileName))
                {
                    return null;
                }

                wb = excel.Workbooks.Open(fileName.Trim());
                Worksheet ws = null;
                //如果有指定工作表名称
                if (!string.IsNullOrEmpty(sheetName))
                {
                    //索引从1开始 //(Excel.Worksheet)wb.Worksheets["SheetName"];
                    ws = (Worksheet)wb.Worksheets[sheetName];
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (ws == null)
                    {
                        ws = (Worksheet)wb.Worksheets[1];
                    }
                }
                else
                {
                    //如果没有指定的sheetName，则尝试获取第一个sheet
                    ws = (Worksheet)wb.Worksheets[1];
                }

                if (ws != null)
                {
                    int rowCount = ws.UsedRange.Rows.Count;//有效行，索引从1开始
                    int columnCount = ws.UsedRange.Columns.Count;//有效列，索引从1开始
                    if (isFirstRowColumn)
                    {
                        //循环列
                        for (int i = 1; i <= columnCount; i++)
                        {
                            if (ws.Cells[0, i].Value2 != null)
                            {
                                System.Data.DataColumn column = new System.Data.DataColumn(ws.Cells[0, i].Value2.ToString());
                                data.Columns.Add(column);
                            }
                        }
                        startRow = 1;
                    }
                    else
                    {
                        startRow = 0;
                    }

                    //循环行
                    for (int i = startRow; i <= rowCount; i++)
                    {
                        System.Data.DataRow dataRow = data.NewRow();
                        //循环列
                        for (int j = 1; j <= columnCount; j++)
                        {
                            if (ws.Cells[i, j].Value2 != null)
                                dataRow[j] = ws.Cells[i, j].Value2.ToString();//取单元格值
                        }
                        data.Rows.Add(dataRow);
                    }
                }
            }
            catch (Exception ex) { 
                MessageBox.Show(ex.Message, "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ClosePro(fileName, excel, wb);
            }
            return data;
        }
        /// <summary>
        /// 写入Excel文件
        /// </summary>
        /// <param name="excelPath">excel 文件路径</param>
        /// <returns></returns>
        public static bool WriteExcel(System.Data.DataTable dt, string path)
        {
            bool result = false;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = null;
            excel.Visible = false;//设置调用引用的 Excel文件是否可见
            excel.DisplayAlerts = false;
            wb = excel.Workbooks.Open(path.Trim());
            Worksheet ws = (Worksheet)wb.Worksheets[1]; //索引从1开始 //(Excel.Worksheet)wb.Worksheets["SheetName"];
            //int rowCount = 0;//有效行，索引从1开始
            try
            {
                int rowCount = dt.Rows.Count;//行数
                int columnCount = dt.Columns.Count;//列数
                if (dt.Rows.Count < ws.UsedRange.Rows.Count)
                    rowCount = ws.UsedRange.Rows.Count;

                //设置每行每列的单元格,  
                for (int i = 0; i < rowCount; i++)
                {
                    if (dt.Rows.Count < ws.UsedRange.Rows.Count && i >= dt.Rows.Count)
                    {
                        for (int j = 0; j < columnCount; j++)
                        {
                            ws.Cells[i + 2, j + 1] = "";
                        }
                    }
                    else {
                        for (int j = 0; j < columnCount; j++)
                        {
                            ws.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                        }
                    }
                }
                result=true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                result = false;
            }
            finally
            {
                ClosePro(path, excel, wb);
            }
            return result;
        }
        /// <summary>
        /// 关闭Excel进程
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="excel"></param>
        /// <param name="wb"></param>
        public static void ClosePro(string path, Microsoft.Office.Interop.Excel.Application excel, Workbook wb)
        {
            Process[] localByNameApp = Process.GetProcessesByName(path);//获取程序名的所有进程
            if (localByNameApp.Length > 0)
            {
                foreach (var app in localByNameApp)
                {
                    if (!app.HasExited)
                    {
                        #region
                        ////设置禁止弹出保存和覆盖的询问提示框   
                        //excel.DisplayAlerts = false;
                        //excel.AlertBeforeOverwriting = false;

                        ////保存工作簿   
                        //excel.Application.Workbooks.Add(true).Save();
                        ////保存excel文件   
                        //excel.Save("D:" + "\\test.xls");
                        ////确保Excel进程关闭   
                        //excel.Quit();
                        //excel = null; 
                        #endregion
                        app.Kill();//关闭进程  
                    }
                }
            }
            if (wb != null)
                wb.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            // 安全回收进程
            System.GC.GetGeneration(excel);
        }
    }
}
