using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;

namespace NCvoucher
{
    public partial class mainForm : Form
    {
        string path1 = "";
        string path2 = "";
        string path3 = "";
        JArray arrList = new JArray();
        JObject detData = new JObject();
        public mainForm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 状态
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_MouseMove(object sender, MouseEventArgs e)
        {
            Point p = e.Location;
            string X = p.X.ToString();
            string Y = p.Y.ToString();
            status.Text = " X:" + X + " Y:" + Y;
        }
        /// <summary>
        /// 鼠标右击菜单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_MouseUp(object sender, MouseEventArgs e)
        {
            string str = dataGridView1.SelectedCells.ToString();
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show(this, e.Location);
            }
        }
        /// <summary>
        /// 快捷键事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            //最大最小窗口
            if (e.KeyCode == Keys.F11)
            {
                if (WindowState == FormWindowState.Maximized)
                    WindowState = FormWindowState.Normal;
                else
                    WindowState = FormWindowState.Maximized;
            }
            //刷新
            if (e.KeyCode == Keys.F5)
            {
                dataGridView1.Rows.Clear();
                //dataGridView1.DataSource = null;
                if (path1 != "")
                {
                    //打开加载框
                    new tipForm(0, "正在加载中...", 1000).Show();

                    dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = dt;
                    //读取Excel
                    arrList = ArrayList.getArrayList(path1);
                    for (int i = 0; i < arrList.Count; i++)
                    {
                        dataGridView1.Rows.Add();
                        int n = 1;
                        foreach (string key in arrList[i])
                        {
                            if (n > 9)
                            {
                                break;
                            }
                            dataGridView1.Rows[i].Cells[n].Value = key;
                            n++;
                        }
                    }
                }
            }
            //保存
            if (e.Modifiers == Keys.Control && e.KeyCode == Keys.S)
            {
                if (path1 == "")
                {
                    MessageBox.Show("请选择文件");
                    return;
                }
                //打开加载框
                new tipForm(0, "正在保存...", 1000).Show();

                DataTable dt = DataGridViewToTable.GetDgvToTable(dataGridView1);
                dt.Columns.RemoveAt(0);
                if (InteropHelper.WriteExcel(dt, path1))
                    //打开加载框
                    new tipForm(1, "保存成功", 1000).Show();
                else
                    new tipForm(2, "保存失败", 1000).Show();
            }
            //打开
            if (e.Modifiers == Keys.Control && e.KeyCode == Keys.O)
            {
                openFileDialog1.InitialDirectory = "C:\\";//初始加载路径为C盘;
                openFileDialog1.Filter = "文本文件 (*.xlsx)|*.xlsx";//过滤你想设置的文本文件类型（这是xml型）
                // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
                if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //打开加载框
                    new tipForm(0, "正在加载中...", 1000).Show();

                    path1 = openFileDialog1.FileName;
                    dataGridView1.Rows.Clear();
                    //dataGridView1.DataSource = dt;
                    //读取Excel
                    arrList = ArrayList.getArrayList(path1);
                    for (int i = 0; i < arrList.Count; i++)
                    {
                        dataGridView1.Rows.Add();
                        int n = 1;
                        foreach (string key in arrList[i])
                        {
                            if (n > 9)
                            {
                                break;
                            }
                            dataGridView1.Rows[i].Cells[n].Value = key;
                            n++;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 打开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void open_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C:\\";//初始加载路径为C盘;
            openFileDialog1.Filter = "文本文件 (*.xlsx)|*.xlsx";//过滤你想设置的文本文件类型（这是xml型）
            // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //打开加载框
                new tipForm(0,"正在加载中...", 1000).Show();

                path1 = openFileDialog1.FileName;
                dataGridView1.Rows.Clear();
                //dataGridView1.DataSource = dt;
                //读取Excel
                arrList=ArrayList.getArrayList(path1);
                for (int i = 0; i < arrList.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    int n = 1;
                    foreach (string key in arrList[i])
                    {
                        if (n > 9)
                        {
                            break;
                        }
                        dataGridView1.Rows[i].Cells[n].Value = key;
                        n++;
                    }
                }
            }
        }
        /// <summary>
        /// 生成凭证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void product_Click(object sender, EventArgs e)
        {
            if (path1 == "")
            {
                MessageBox.Show("请选择文件");
                return;
            }

            detData["company"] = company.Text;
            detData["voucher_type"] = voucher_type.Text;
            detData["prepareddate"] = prepareddate.Text;
            detData["enter"] = enter.Text;
            detData["Val1"] = Val1.Text;
            detData["Val2"] = Val2.Text;
            detData["Val3"] = Val3.Text;
            detData["Val4"] = Val4.Text;
            detData["Val5"] = Val5.Text;

            //打开加载框
            new tipForm(0,"凭证生成中...", 1000).Show();

            //生成XML
            string Path = AppDomain.CurrentDomain.BaseDirectory + "test.xml";
            OptionFile.DeleteFile(Path);
            //OptionFile.CreateFile(Path);
            XmlDocument document = new XmlDocument();
            document.LoadXml(ProduceXml.GetXml(arrList, detData));
            document.Save(Path);
            //MessageBox.Show("凭证已生成");
        }
        /// <summary>
        /// 下载凭证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void download_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = "C:\\";//初始加载路径为C盘;
            saveFileDialog1.FileName = "voucher_u8.xml";
            saveFileDialog1.Filter = "文本文件 (*.xml)|*.xml";//过滤你想设置的文本文件类型（这是xml型）
            // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //打开加载框
                new tipForm(0,"正在下载...", 1000).Show();

                path2 = saveFileDialog1.FileName;
                if (!File.Exists(path2))
                {
                    FileStream stream = File.Create(path2);
                    stream.Close();
                    stream.Dispose();
                }
                File.Copy(AppDomain.CurrentDomain.BaseDirectory + "test.xml", path2, true);
                //MessageBox.Show("下载成功");
            }
        }
        /// <summary>
        /// 查看凭证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void look_Click(object sender, EventArgs e)
        {
            FormDialog formdialog = new FormDialog();
            formdialog.Show();
        }
        /// <summary>
        /// 全选
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selall_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //设置每一行的选择框为选中
                dataGridView1.Rows[i].Cells[0].Value = true;
            }
        }
        /// <summary>
        /// 取消
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancel_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //设置每一行的选择框为未选中
                dataGridView1.Rows[i].Cells[0].Value = false;
            }
        }
        /// <summary>
        /// 反选
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void invertsel_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //判断当前行是否被选中
                if ((bool)dataGridView1.Rows[i].Cells[0].EditedFormattedValue == true)
                    //设置每一行的选择框为未选中
                    dataGridView1.Rows[i].Cells[0].Value = false;
                else
                    //设置每一行的选择框为选中
                    dataGridView1.Rows[i].Cells[0].Value = true;
            }
        }
        /// <summary>
        /// 刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void refresh_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            //dataGridView1.DataSource = null;
            if (path1 != "") {
                //打开加载框
                new tipForm(0,"正在加载中...", 1000).Show();
                
                dataGridView1.Rows.Clear();
                //dataGridView1.DataSource = dt;
                //读取Excel
                arrList = ArrayList.getArrayList(path1);
                for (int i = 0; i < arrList.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    int n = 1;
                    foreach (string key in arrList[i])
                    {
                        if (n > 9)
                        {
                            break;
                        }
                        dataGridView1.Rows[i].Cells[n].Value = key;
                        n++;
                    }
                }
            }
        }
        /// <summary>
        /// 清空
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void clean_Click(object sender, EventArgs e)
        {
            //打开加载框
            new tipForm(0, "正在清空...", 1000).Show();

            path1 = "";
            dataGridView1.Rows.Clear();
            //dataGridView1.DataSource = null;
        }
        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void save_Click(object sender, EventArgs e)
        {
            if (path1 == "") {
                MessageBox.Show("请选择文件");
                return;
            }
            //打开加载框
            new tipForm(0, "正在保存...", 1000).Show();

            DataTable dt =DataGridViewToTable.GetDgvToTable(dataGridView1);
            dt.Columns.RemoveAt(0);
            if (InteropHelper.WriteExcel(dt, path1))
                //打开加载框
                new tipForm(1, "保存成功", 1000).Show();
            else
                new tipForm(2, "保存失败", 1000).Show();
        }
        /// <summary>
        /// 另存为
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveas_Click(object sender, EventArgs e)
        {
            saveFileDialog2.InitialDirectory = "C:\\";//初始加载路径为C盘;
            saveFileDialog2.FileName = "excel.xlsx";
            saveFileDialog2.Filter = "文本文件 (*.xlsx)|*.xlsx";//过滤你想设置的文本文件类型（这是xml型）
            // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
            if (this.saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                //打开加载框
                new tipForm(0, "正在保存...", 1000).Show();

                path3 = saveFileDialog2.FileName;
                if (!File.Exists(path3))
                {
                    FileStream stream = File.Create(path3);
                    stream.Close();
                    stream.Dispose();
                }
                File.Copy(path1, path3, true);
                DataTable dt = DataGridViewToTable.GetDgvToTable(dataGridView1);
                dt.Columns.RemoveAt(0);
                if (InteropHelper.WriteExcel(dt, path3))
                    //打开加载框
                    new tipForm(1, "保存成功", 1000).Show();
                else
                    //打开加载框
                    new tipForm(2, "保存失败", 1000).Show();

                //DataTable dt =DataGridViewToTable.GetDgvToTable(dataGridView1);
                //dt.Columns.RemoveAt(0);
                //if (NPOIHelper.WriteDataTableToExcel(dt, path3))
                //    MessageBox.Show("保存成功");
                //else
                //    MessageBox.Show("保存失败");
            }
        }
        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void delete_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //判断当前行是否被选中
                if ((bool)dataGridView1.Rows[i].Cells[0].EditedFormattedValue == true)
                {
                    try
                    {
                        dataGridView1.Rows.RemoveAt(i);
                        i--;
                    }
                    catch { }
                }
            }
            //打开加载框
            new tipForm(1, "删除成功", 1000).Show();
            //MessageBox.Show("删除成功");
        }
        /// <summary>
        /// 字体
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void font_Click(object sender, EventArgs e)
        {
            // 打开字体对话框
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                dataGridView1.Font = fontDialog1.Font;
            }
        }
        /// <summary>
        /// 风格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void color_Click(object sender, EventArgs e)
        {
            // 打开颜色对话框
            if (colorDialog1.ShowDialog() != DialogResult.Cancel)
            {
                this.BackColor = colorDialog1.Color;
                menuStrip1.BackColor = colorDialog1.Color;
            }
        }
        /// <summary>
        /// 关于
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void about_Click(object sender, EventArgs e)
        {
            aboutForm aboutform = new aboutForm();
            aboutform.Show();
        }

        private void right_open_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C:\\";//初始加载路径为C盘;
            openFileDialog1.Filter = "文本文件 (*.xlsx)|*.xlsx";//过滤你想设置的文本文件类型（这是xml型）
            // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //打开加载框
                new tipForm(0, "正在加载中...", 1000).Show();

                path1 = openFileDialog1.FileName;
                dataGridView1.Rows.Clear();
                //dataGridView1.DataSource = dt;
                //读取Excel
                arrList = ArrayList.getArrayList(path1);
                for (int i = 0; i < arrList.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    int n = 1;
                    foreach (string key in arrList[i])
                    {
                        if (n > 9)
                        {
                            break;
                        }
                        dataGridView1.Rows[i].Cells[n].Value = key;
                        n++;
                    }
                }
            }
        }

        private void right_save_Click(object sender, EventArgs e)
        {
            if (path1 == "")
            {
                MessageBox.Show("请选择文件");
                return;
            }
            //打开加载框
            new tipForm(0, "正在保存...", 1000).Show();

            DataTable dt = DataGridViewToTable.GetDgvToTable(dataGridView1);
            dt.Columns.RemoveAt(0);
            if (InteropHelper.WriteExcel(dt, path1))
                //打开加载框
                new tipForm(1, "保存成功", 1000).Show();
            else
                new tipForm(2, "保存失败", 1000).Show();
        }

        private void right_saveas_Click(object sender, EventArgs e)
        {
            saveFileDialog2.InitialDirectory = "C:\\";//初始加载路径为C盘;
            saveFileDialog2.FileName = "excel.xlsx";
            saveFileDialog2.Filter = "文本文件 (*.xlsx)|*.xlsx";//过滤你想设置的文本文件类型（这是xml型）
            // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
            if (this.saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                //打开加载框
                new tipForm(0, "正在保存...", 1000).Show();

                path3 = saveFileDialog2.FileName;
                if (!File.Exists(path3))
                {
                    FileStream stream = File.Create(path3);
                    stream.Close();
                    stream.Dispose();
                }
                File.Copy(path1, path3, true);
                DataTable dt = DataGridViewToTable.GetDgvToTable(dataGridView1);
                dt.Columns.RemoveAt(0);
                if (InteropHelper.WriteExcel(dt, path3))
                    //打开加载框
                    new tipForm(1, "保存成功", 1000).Show();
                else
                    //打开加载框
                    new tipForm(2, "保存失败", 1000).Show();
            }
        }

        private void right_download_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = "C:\\";//初始加载路径为C盘;
            saveFileDialog1.FileName = "voucher_u8.xml";
            saveFileDialog1.Filter = "文本文件 (*.xml)|*.xml";//过滤你想设置的文本文件类型（这是xml型）
            // openFileDialog1.Filter = "文本文件 (*.txt)|*.txt|All files (*.*)|*.*";（这是全部类型文件）
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //打开加载框
                new tipForm(0, "正在下载...", 1000).Show();

                path2 = saveFileDialog1.FileName;
                if (!File.Exists(path2))
                {
                    FileStream stream = File.Create(path2);
                    stream.Close();
                    stream.Dispose();
                }
                File.Copy(AppDomain.CurrentDomain.BaseDirectory + "test.xml", path2, true);
            }
        }

        private void right_product_Click(object sender, EventArgs e)
        {
            if (path1 == "")
            {
                MessageBox.Show("请选择文件");
                return;
            }

            detData["company"] = company.Text;
            detData["voucher_type"] = voucher_type.Text;
            detData["prepareddate"] = prepareddate.Text;
            detData["enter"] = enter.Text;
            detData["Val1"] = Val1.Text;
            detData["Val2"] = Val2.Text;
            detData["Val3"] = Val3.Text;
            detData["Val4"] = Val4.Text;
            detData["Val5"] = Val5.Text;

            //打开加载框
            new tipForm(0, "凭证生成中...", 1000).Show();

            //生成XML
            string Path = AppDomain.CurrentDomain.BaseDirectory + "test.xml";
            OptionFile.DeleteFile(Path);
            //OptionFile.CreateFile(Path);
            XmlDocument document = new XmlDocument();
            document.LoadXml(ProduceXml.GetXml(arrList, detData));
            document.Save(Path);
        }

        private void right_selall_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //设置每一行的选择框为选中
                dataGridView1.Rows[i].Cells[0].Value = true;
            }
        }

        private void right_invertsel_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //判断当前行是否被选中
                if ((bool)dataGridView1.Rows[i].Cells[0].EditedFormattedValue == true)
                    //设置每一行的选择框为未选中
                    dataGridView1.Rows[i].Cells[0].Value = false;
                else
                    //设置每一行的选择框为选中
                    dataGridView1.Rows[i].Cells[0].Value = true;
            }
        }

        private void right_cancel_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //设置每一行的选择框为未选中
                dataGridView1.Rows[i].Cells[0].Value = false;
            }
        }

        private void right_delete_Click(object sender, EventArgs e)
        {
            //循环dataGridView
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //判断当前行是否被选中
                if ((bool)dataGridView1.Rows[i].Cells[0].EditedFormattedValue == true)
                {
                    try
                    {
                        dataGridView1.Rows.RemoveAt(i);
                        i--;
                    }
                    catch { }
                }
            }
            //打开加载框
            new tipForm(1, "删除成功", 1000).Show();
        }

        private void right_clean_Click(object sender, EventArgs e)
        {
            //打开加载框
            new tipForm(0, "正在清空...", 1000).Show();

            path1 = "";
            dataGridView1.Rows.Clear();
            //dataGridView1.DataSource = null;
        }

        private void right_look_Click(object sender, EventArgs e)
        {
            FormDialog formdialog = new FormDialog();
            formdialog.Show();
        }

        private void right_about_Click(object sender, EventArgs e)
        {
            aboutForm aboutform = new aboutForm();
            aboutform.Show();
        }

        private void right_font_Click(object sender, EventArgs e)
        {
            // 打开字体对话框
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                dataGridView1.Font = fontDialog1.Font;
            }
        }

        private void right_color_Click(object sender, EventArgs e)
        {
            // 打开颜色对话框
            if (colorDialog1.ShowDialog() != DialogResult.Cancel)
            {
                this.BackColor = colorDialog1.Color;
                menuStrip1.BackColor = colorDialog1.Color;
            }
        }

        private void right_refresh_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            //dataGridView1.DataSource = null;
            if (path1 != "")
            {
                //打开加载框
                new tipForm(0, "正在加载中...", 1000).Show();

                dataGridView1.Rows.Clear();
                //dataGridView1.DataSource = dt;
                //读取Excel
                arrList = ArrayList.getArrayList(path1);
                for (int i = 0; i < arrList.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    int n = 1;
                    foreach (string key in arrList[i])
                    {
                        if (n > 9)
                        {
                            break;
                        }
                        dataGridView1.Rows[i].Cells[n].Value = key;
                        n++;
                    }
                }
            }
        }
    }
}
