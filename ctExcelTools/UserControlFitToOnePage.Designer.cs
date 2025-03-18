using System.Windows.Forms;

namespace ctExcelTools
{
	partial class UserControlFitToOnePage
	{
		/// <summary> 
		/// 必需的设计器变量。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// 清理所有正在使用的资源。
		/// </summary>
		/// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
		protected override void Dispose(bool disposing)
		{
			//MessageBox.Show("UserControlFitToOnePage Dispose");

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
			this.textBox_Info = new System.Windows.Forms.TextBox();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.SuspendLayout();
			// 
			// textBox_Info
			// 
			this.textBox_Info.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.textBox_Info.HideSelection = false;
			this.textBox_Info.Location = new System.Drawing.Point(3, 3);
			this.textBox_Info.Multiline = true;
			this.textBox_Info.Name = "textBox_Info";
			this.textBox_Info.ReadOnly = true;
			this.textBox_Info.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.textBox_Info.Size = new System.Drawing.Size(147, 147);
			this.textBox_Info.TabIndex = 0;
			this.textBox_Info.WordWrap = false;
			// 
			// timer1
			// 
			this.timer1.Enabled = true;
			this.timer1.Interval = 1000;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// UserControlFitToOnePage
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.textBox_Info);
			this.Name = "UserControlFitToOnePage";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		public System.Windows.Forms.TextBox textBox_Info;
		private Timer timer1;
	}
}
