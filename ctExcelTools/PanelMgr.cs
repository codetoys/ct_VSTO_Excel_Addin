using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Windows.Forms;

namespace ctExcelTools
{
	//管理自定义任务窗格（面板），注意CustomTaskPanes.Add会在当前活动工作簿上创建任务窗格
	public class PanelMgr
	{
		private Dictionary<Microsoft.Office.Interop.Excel.Workbook, (CustomTaskPane, UserControlFitToOnePage)> m_Panels = new Dictionary<Microsoft.Office.Interop.Excel.Workbook, (CustomTaskPane, UserControlFitToOnePage)>();
		//警告，不可在工作者线程调用，对于工作者线程的workbook将不认为是同一个对象，m_Panels.ContainsKey(workbook)将返回false
		//同时“UserControlFitToOnePage userControlFitToOnePage = new UserControlFitToOnePage();”这一句将引发异常“不支持的接口”
		public (CustomTaskPane, UserControlFitToOnePage) GetPanel(Microsoft.Office.Interop.Excel.Workbook workbook)
		{
			if (!m_Panels.ContainsKey(workbook))
			{
				//MessageBox.Show("创建", workbook.Name);
				UserControlFitToOnePage userControlFitToOnePage = new UserControlFitToOnePage();
				CustomTaskPane panel = Globals.ThisAddIn.CustomTaskPanes.Add(userControlFitToOnePage, workbook.Name);
				m_Panels.Add(workbook, (panel, userControlFitToOnePage));
			}
			if (!m_Panels.ContainsKey(workbook))
			{
				MessageBox.Show("不应该", workbook.Name);
			}
			return m_Panels[workbook];
		}
		public void Remove(Microsoft.Office.Interop.Excel.Workbook workbook)
		{
			m_Panels.Remove(workbook);
		}
	}
}
