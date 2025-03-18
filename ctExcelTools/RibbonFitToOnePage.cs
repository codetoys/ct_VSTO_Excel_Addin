using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ctExcelTools
{
	public partial class RibbonFitToOnePage
	{
		private void RibbonFitToOnePage_Load(object sender, RibbonUIEventArgs e)
		{
			Log.LogI("RibbonFitToOnePage_Load");
		}
		private void ShowlogWindow()
		{
			Globals.ThisAddIn.Form_Log.Show();
			Globals.ThisAddIn.Form_Log.TopMost = true;//调到前端显示
			Globals.ThisAddIn.Form_Log.TopMost = false;//恢复正常，否则始终在最前端
		}
		private Range GetPrintRange(Worksheet worksheet)
		{
			if (null != worksheet.PageSetup.PrintArea && worksheet.PageSetup.PrintArea.Length > 0) return worksheet.get_Range(worksheet.PageSetup.PrintArea);
			else return worksheet.UsedRange;
		}
		private void button_ShowInfo_Click(object sender, RibbonControlEventArgs e)
		{
			ShowlogWindow();

			Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
			Log.InfoOutput(workbook, "开始操作。。。。。。");
			try
			{
				Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
				Log.InfoOutput(workbook, " UsedRange：" + worksheet.UsedRange.Address);
				Log.InfoOutput(workbook, " PrintArea：" + worksheet.PageSetup.PrintArea);
				Log.InfoOutput(workbook, " PaperSize：" + worksheet.PageSetup.PaperSize.ToString());
				Log.InfoOutput(workbook, " ChartSize：" + worksheet.PageSetup.ChartSize);
				Log.InfoOutput(workbook, " Orientation：" + worksheet.PageSetup.Orientation.ToString());
				Log.InfoOutput(workbook, " TopMargin：" + worksheet.PageSetup.TopMargin);
				Log.InfoOutput(workbook, " BottomMargin：" + worksheet.PageSetup.BottomMargin);
				Log.InfoOutput(workbook, " LeftMargin：" + worksheet.PageSetup.LeftMargin);
				Log.InfoOutput(workbook, " RightMargin：" + worksheet.PageSetup.RightMargin);
				Log.InfoOutput(workbook, " Pages：" + worksheet.PageSetup.Pages.Count);
				Log.InfoOutput(workbook, " Zoom：" + worksheet.PageSetup.Zoom);
				Range printRange = GetPrintRange(worksheet);
				Log.InfoOutput(workbook, " printRange.Column：" + printRange.Column);
				Log.InfoOutput(workbook, " printRange.Columns.Count：" + printRange.Columns.Count);
				Log.InfoOutput(workbook, " printRange.Row：" + printRange.Row);
				Log.InfoOutput(workbook, " printRange.Rows.Count：" + printRange.Rows.Count);

				double originalTotalWidth = 0;
				double originalTotalHeigh = 0;

				for (int i = 0; i < printRange.Columns.Count; ++i)
				{
					Range colum = worksheet.Columns[printRange.Column + i];
					originalTotalWidth += colum.ColumnWidth;
				}
				for (int i = 0; i < printRange.Rows.Count; ++i)
				{
					Range row = worksheet.Rows[printRange.Row + i];
					originalTotalHeigh += row.RowHeight;
				}
				Log.InfoOutput(workbook, " originalTotalWidth：" + originalTotalWidth);
				Log.InfoOutput(workbook, " originalTotalHeigh：" + originalTotalHeigh);

				Log.InfoOutput(workbook, "操作成功完成");
			}
			catch (Exception ex)
			{
				Log.InfoOutput(workbook, ex.ToString());
			}
		}

		private void button_FitToOnePage_Click(object sender, RibbonControlEventArgs e)
		{
			string str = "开始操作。。。。。。\n";
			try
			{
				Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
				if (worksheet.PageSetup.Zoom.GetType() == false.GetType() && worksheet.PageSetup.Zoom == false)//动态类型，可能是数字或bool
				{
					if (MessageBox.Show("当前是自动调整模式，必须以缩放模式调整，将设置为100%缩放模式（无缩放）继续吗？", "当前缩放" + worksheet.PageSetup.Zoom, MessageBoxButtons.OKCancel) != DialogResult.OK)
					{
						return;
					}
					worksheet.PageSetup.Zoom = 100;
				}
				if (worksheet.PageSetup.Pages.Count > 1)
				{
					MessageBox.Show("当前为" + worksheet.PageSetup.Pages.Count + "页，请先想办法缩小至一页");
					return;
				}
				Range printRange = GetPrintRange(worksheet);
				int firstColumn = printRange.Column;
				int columnCount = printRange.Columns.Count;
				int firstRow = printRange.Row;
				int rowCount = printRange.Rows.Count;

				str += " Pages：" + worksheet.PageSetup.Pages.Count + "\n";

				double originalTotalWidth = 0;
				double originalTotalHeigh = 0;
				double[] originalWidth = new double[columnCount];
				double[] originalHeight = new double[rowCount];

				for (int i = 0; i < columnCount; ++i)
				{
					Range colum = worksheet.Columns[firstColumn + i];
					originalWidth[i] = colum.ColumnWidth;
					originalTotalWidth += colum.ColumnWidth;
				}
				for (int i = 0; i < rowCount; ++i)
				{
					Range row = worksheet.Rows[firstRow + i];
					originalHeight[i] = row.RowHeight;
					originalTotalHeigh += row.RowHeight;
				}
				str += " originalTotalWidth：" + originalTotalWidth + "\n";
				str += " originalTotalHeigh：" + originalTotalHeigh + "\n";

				double step = 100;//步进，如果超出一页就减小步进值，直到步进小于某个值
				double fix = 0;

				while (worksheet.PageSetup.Pages.Count == 1 && step >= 0.1)//注意列宽单位是标准字符
				{
					fix += step;

					for (int i = 0; i < columnCount; ++i)
					{
						Range colum = worksheet.Columns[firstColumn + i];
						colum.ColumnWidth = originalWidth[i] + fix * 0.1;
					}
					if (worksheet.PageSetup.Pages.Count > 1)
					{
						fix -= step;
						step /= 2;
						for (int i = 0; i < columnCount; ++i)
						{
							Range colum = worksheet.Columns[firstColumn + i];
							colum.ColumnWidth = originalWidth[i] + fix * 0.1;
						}
					}
				}
				str += " Pages：" + worksheet.PageSetup.Pages.Count + "\n";
				step = 100;
				fix = 0;
				while (worksheet.PageSetup.Pages.Count == 1 && step >= 1)//注意行高单位是像素
				{
					fix += step;

					for (int i = 0; i < rowCount; ++i)
					{
						Range row = worksheet.Rows[firstRow + i];
						row.RowHeight = originalHeight[i] + fix;
					}
					if (worksheet.PageSetup.Pages.Count > 1)
					{
						fix -= step;
						step /= 2;
						for (int i = 0; i < rowCount; ++i)
						{
							Range row = worksheet.Rows[firstRow + i];
							row.RowHeight = originalHeight[i] + fix;
						}
					}
				}

				str += " Pages：" + worksheet.PageSetup.Pages.Count + "\n";
				str += "操作成功完成\n";
				worksheet.PrintPreview();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
			MessageBox.Show(str);
		}

		private void button_AutoFit_Click(object sender, RibbonControlEventArgs e)
		{
			string str = "开始操作。。。。。。\n";
			try
			{
				Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
				Range printRange = GetPrintRange(worksheet);

				for (int i = 0; i < printRange.Columns.Count; ++i)
				{
					Range colum = worksheet.Columns[printRange.Column + i];
					colum.AutoFit();
				}
				for (int i = 0; i < printRange.Rows.Count; ++i)
				{
					Range row = worksheet.Rows[printRange.Row + i];
					row.AutoFit();
				}
				str += " Pages：" + worksheet.PageSetup.Pages.Count + "\n";
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
			MessageBox.Show(str);
		}

		private void changeFontSize(int n)
		{
			try
			{
				Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
				Range printRange = GetPrintRange(worksheet);

				for (int row = 0; row < printRange.Rows.Count; ++row)
				{
					for (int col = 0; col < printRange.Columns.Count; ++col)
					{
						Range cell = worksheet.Cells[printRange.Column + col, printRange.Row + row];
						cell.Font.Size += n;
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}
		private void button_FontInc_Click(object sender, RibbonControlEventArgs e)
		{
			changeFontSize(1);
		}

		private void button_FontDec_Click(object sender, RibbonControlEventArgs e)
		{
			changeFontSize(-1);
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			//Log.InfoOutput(Globals.ThisAddIn.Application.ActiveWorkbook,DateTime.Now.ToString()+ "功能区定时器");
		}

		private void button_ShowLogWindow_Click(object sender, RibbonControlEventArgs e)
		{
			ShowlogWindow();
		}
	}
}
