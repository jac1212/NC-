using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NCvoucher
{
    public partial class FormDialog : Form
    {
        public FormDialog()
        {
            InitializeComponent();
            string line;
            string str = "";
            try
            {
                StreamReader sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "test.xml", Encoding.Default);
                while ((line = sr.ReadLine()) != null)
                {
                    str += line + "\r\n";
                }
                sr.Close();
                textBox1.Text = str;
            }
            catch (IOException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
