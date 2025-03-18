using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ctExcelTools
{
    static class Log
    {
		public static void LogI(string str)
		{
			Globals.ThisAddIn.Form_Log.textBox_Info.Text += DateTime.Now.ToString() + "[信息]" + str + "\r\n";
		}
		public static void InfoOutput(Workbook workbook, string str)
		{
			LogI("[" + workbook.Name + "]" + str);
			//暂时不用任务窗格，但相关代码是经过验证的
			//(CustomTaskPane, UserControlFitToOnePage) tmp = Globals.ThisAddIn.panelMgr.GetPanel(workbook);
			//tmp.Item1.Visible = true;
			//tmp.Item2.textBox_Info.Text += DateTime.Now.ToString() + "[信息]" + str + "\r\n";
		}
	}
}
