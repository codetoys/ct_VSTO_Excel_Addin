using System.Windows.Forms;

namespace ctExcelTools
{
	partial class RibbonFitToOnePage : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// 必需的设计器变量。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public RibbonFitToOnePage()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// 清理所有正在使用的资源。
		/// </summary>
		/// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
		protected override void Dispose(bool disposing)
		{
			//MessageBox.Show("RibbonFitToOnePage Dispose");
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region 组件设计器生成的代码

		/// <summary>
		/// 设计器支持所需的方法 - 不要修改
		/// 使用代码编辑器修改此方法的内容。
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.button_ShowState = this.Factory.CreateRibbonButton();
			this.button_FontInc = this.Factory.CreateRibbonButton();
			this.button_FontDec = this.Factory.CreateRibbonButton();
			this.button_AutoFit = this.Factory.CreateRibbonButton();
			this.button_FitToOnePage = this.Factory.CreateRibbonButton();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.button_ShowLogWindow = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.group1);
			this.tab1.Label = "TabAddIns";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.button_ShowLogWindow);
			this.group1.Items.Add(this.button_ShowState);
			this.group1.Items.Add(this.button_FontInc);
			this.group1.Items.Add(this.button_FontDec);
			this.group1.Items.Add(this.button_AutoFit);
			this.group1.Items.Add(this.button_FitToOnePage);
			this.group1.Label = "调整到一页";
			this.group1.Name = "group1";
			// 
			// button_ShowState
			// 
			this.button_ShowState.Label = "信息";
			this.button_ShowState.Name = "button_ShowState";
			this.button_ShowState.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ShowInfo_Click);
			// 
			// button_FontInc
			// 
			this.button_FontInc.Label = "字号+";
			this.button_FontInc.Name = "button_FontInc";
			this.button_FontInc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_FontInc_Click);
			// 
			// button_FontDec
			// 
			this.button_FontDec.Label = "字号-";
			this.button_FontDec.Name = "button_FontDec";
			this.button_FontDec.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_FontDec_Click);
			// 
			// button_AutoFit
			// 
			this.button_AutoFit.Label = "AutoFit";
			this.button_AutoFit.Name = "button_AutoFit";
			this.button_AutoFit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_AutoFit_Click);
			// 
			// button_FitToOnePage
			// 
			this.button_FitToOnePage.Label = "FitToOnePage";
			this.button_FitToOnePage.Name = "button_FitToOnePage";
			this.button_FitToOnePage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_FitToOnePage_Click);
			// 
			// timer1
			// 
			this.timer1.Enabled = true;
			this.timer1.Interval = 1000;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// button_ShowLogWindow
			// 
			this.button_ShowLogWindow.Label = "日志";
			this.button_ShowLogWindow.Name = "button_ShowLogWindow";
			this.button_ShowLogWindow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ShowLogWindow_Click);
			// 
			// RibbonFitToOnePage
			// 
			this.Name = "RibbonFitToOnePage";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonFitToOnePage_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button_ShowState;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button_FitToOnePage;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button_AutoFit;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button_FontInc;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button_FontDec;
		private Timer timer1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button_ShowLogWindow;
	}

	partial class ThisRibbonCollection
	{
		internal RibbonFitToOnePage RibbonFitToOnePage
		{
			get { return this.GetRibbon<RibbonFitToOnePage>(); }
		}
	}
}
