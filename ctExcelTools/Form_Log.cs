using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ctExcelTools
{
    public partial class Form_Log: Form
    {
		public delegate void delegateAddInfo(string message);
		public delegateAddInfo AddInfo;
		public void _AddInfo(string message)
		{
			Log.InfoOutput(Globals.ThisAddIn.Application.ActiveWorkbook, message);
		}
		public void CallAddInfo(string message)
		{
			this.Invoke(AddInfo, message);
		}
		public Form_Log()
        {
            InitializeComponent();
	
			AddInfo = new delegateAddInfo(_AddInfo);
		}

		private void Form_Log_FormClosing(object sender, FormClosingEventArgs e)
		{
            e.Cancel = true;
            this.Hide();
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
			{
				//Log.InfoOutput(Globals.ThisAddIn.Application.ActiveWorkbook, DateTime.Now.ToString() + " 独立窗体定时器");
			}
		}
	}
}
