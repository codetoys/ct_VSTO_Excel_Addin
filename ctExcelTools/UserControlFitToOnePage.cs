using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ctExcelTools
{
    public partial class UserControlFitToOnePage: UserControl
    {
		public UserControlFitToOnePage()
        {
            InitializeComponent();
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			//this.textBox_Info.Text += DateTime.Now.ToString() + " 用户控件定时器\r\n";
		}
	}
}
