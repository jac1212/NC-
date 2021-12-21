using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NCvoucher
{
    public partial class tipForm : Form
    {
        public tipForm(int icon,string msg,int ts)
        {
            InitializeComponent();
            timer1.Enabled = true;
            timer1.Interval = ts;
            tip.Text=msg;
            switch (icon) { 
                case 1:
                    pic1.Hide();
                    pic2.ImageLocation = "icon/success.ico";
                    break;
                case 2:
                    pic1.Hide();
                    pic2.ImageLocation = "icon/error.ico";
                    break;
                default:
                    pic1.ImageLocation = "icon/loading.gif";
                    pic2.Hide();
                    break;
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
