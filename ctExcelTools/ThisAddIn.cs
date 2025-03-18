using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using static System.Net.Mime.MediaTypeNames;

namespace ctExcelTools
{
	public partial class ThisAddIn
	{
		public Form_Log Form_Log;
		public PanelMgr panelMgr = new PanelMgr();
		bool bExit = false;
		Thread thread;
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			Form_Log = new Form_Log();

			this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
			this.Application.WorkbookOpen += Application_WorkbookOpen;
			this.Application.SheetChange += Application_SheetChange;

			thread = new Thread(thread_OnTime);
			//thread.SetApartmentState(System.Threading.ApartmentState.STA);//似乎不是必须
			//thread.Start();

			Log.LogI("ThisAddIn_Startup");
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			//MessageBox.Show("ThisAddIn_Shutdown");
			bExit = true;
			thread.Join();
		}
		private void Application_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook workbook, ref bool Cancel)
		{
			//MessageBox.Show("Application_WorkbookBeforeClose");
			panelMgr.Remove(workbook);
		}
		private void Application_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook workbook)
		{
			//MessageBox.Show("Application_WorkbookOpen");
		}
		private void Application_SheetChange(object Sh, Range Target)
		{
			//MessageBox.Show("Application_SheetChange " + Target.Worksheet.Name + " " + Target.Address);
		}
		private void thread_OnTime()
		{
			DateTime dateTime = DateTime.Now;
			//MessageBox.Show("1");
			while (!bExit)
			{
				try
				{
					if (null != this.Application.ActiveWorkbook && (DateTime.Now - dateTime).TotalMilliseconds >= 5000)//this.Application.ActiveWorkbook可能在退出时引发异常
					{
						Form_Log.CallAddInfo(DateTime.Now.ToString() + DateTime.Now.ToString() + "工作者线程");
						dateTime = DateTime.Now;
						//MessageBox.Show(dateTime.ToString());
					}
					else Thread.Yield();
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.ToString());
					return;
				}
			}
			//MessageBox.Show("2");
		}

		#region VSTO 生成的代码

		/// <summary>
		/// 设计器支持所需的方法 - 不要修改
		/// 使用代码编辑器修改此方法的内容。
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
