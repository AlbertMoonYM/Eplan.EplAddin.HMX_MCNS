namespace Eplan.EplAddin.HMX_MCNS._03.UI
{
    partial class FormIoList
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlTap = new DevExpress.XtraEditors.PanelControl();
            this.lblLogo = new DevExpress.XtraEditors.LabelControl();
            this.picBoxLogo = new DevExpress.XtraEditors.PictureEdit();
            this.panelControl4 = new DevExpress.XtraEditors.PanelControl();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.panelControl5 = new DevExpress.XtraEditors.PanelControl();
            this.btnSaveIo = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.pnlTap)).BeginInit();
            this.pnlTap.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxLogo.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl4)).BeginInit();
            this.panelControl4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl5)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlTap
            // 
            this.pnlTap.Controls.Add(this.lblLogo);
            this.pnlTap.Controls.Add(this.picBoxLogo);
            this.pnlTap.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTap.Location = new System.Drawing.Point(0, 0);
            this.pnlTap.Name = "pnlTap";
            this.pnlTap.Size = new System.Drawing.Size(1084, 60);
            this.pnlTap.TabIndex = 32;
            // 
            // lblLogo
            // 
            this.lblLogo.Appearance.Font = new System.Drawing.Font("Microsoft YaHei UI", 20F, System.Drawing.FontStyle.Bold);
            this.lblLogo.Appearance.Options.UseFont = true;
            this.lblLogo.Location = new System.Drawing.Point(66, 14);
            this.lblLogo.Name = "lblLogo";
            this.lblLogo.Size = new System.Drawing.Size(96, 36);
            this.lblLogo.TabIndex = 32;
            this.lblLogo.Text = "IO LIST";
            // 
            // picBoxLogo
            // 
            this.picBoxLogo.EditValue = global::Eplan.EplAddin.HMX_MCNS.Properties.Resources.IO;
            this.picBoxLogo.Location = new System.Drawing.Point(10, 5);
            this.picBoxLogo.Name = "picBoxLogo";
            this.picBoxLogo.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.picBoxLogo.Properties.Appearance.Options.UseBackColor = true;
            this.picBoxLogo.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.picBoxLogo.Properties.ShowCameraMenuItem = DevExpress.XtraEditors.Controls.CameraMenuItemVisibility.Auto;
            this.picBoxLogo.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Stretch;
            this.picBoxLogo.Size = new System.Drawing.Size(52, 52);
            this.picBoxLogo.TabIndex = 31;
            // 
            // panelControl4
            // 
            this.panelControl4.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panelControl4.Controls.Add(this.gridControl1);
            this.panelControl4.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControl4.Location = new System.Drawing.Point(0, 65);
            this.panelControl4.Name = "panelControl4";
            this.panelControl4.Size = new System.Drawing.Size(1084, 650);
            this.panelControl4.TabIndex = 76;
            // 
            // gridControl1
            // 
            this.gridControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridControl1.Location = new System.Drawing.Point(0, 0);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(1084, 650);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.Name = "gridView1";
            // 
            // panelControl5
            // 
            this.panelControl5.Appearance.BackColor = System.Drawing.Color.DimGray;
            this.panelControl5.Appearance.Options.UseBackColor = true;
            this.panelControl5.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panelControl5.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControl5.Location = new System.Drawing.Point(0, 60);
            this.panelControl5.Name = "panelControl5";
            this.panelControl5.Size = new System.Drawing.Size(1084, 5);
            this.panelControl5.TabIndex = 75;
            // 
            // btnSaveIo
            // 
            this.btnSaveIo.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnSaveIo.Appearance.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnSaveIo.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9.75F, System.Drawing.FontStyle.Bold);
            this.btnSaveIo.Appearance.Options.UseBackColor = true;
            this.btnSaveIo.Appearance.Options.UseFont = true;
            this.btnSaveIo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSaveIo.Location = new System.Drawing.Point(475, 721);
            this.btnSaveIo.Name = "btnSaveIo";
            this.btnSaveIo.Size = new System.Drawing.Size(125, 38);
            this.btnSaveIo.TabIndex = 77;
            this.btnSaveIo.Text = "저장하기";
            // 
            // FormIoList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1084, 784);
            this.Controls.Add(this.btnSaveIo);
            this.Controls.Add(this.panelControl4);
            this.Controls.Add(this.panelControl5);
            this.Controls.Add(this.pnlTap);
            this.IconOptions.Image = global::Eplan.EplAddin.HMX_MCNS.Properties.Resources.MCNS_LOGO;
            this.Name = "FormIoList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "IO LIST";
            ((System.ComponentModel.ISupportInitialize)(this.pnlTap)).EndInit();
            this.pnlTap.ResumeLayout(false);
            this.pnlTap.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxLogo.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl4)).EndInit();
            this.panelControl4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl5)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.PanelControl pnlTap;
        private DevExpress.XtraEditors.LabelControl lblLogo;
        private DevExpress.XtraEditors.PictureEdit picBoxLogo;
        private DevExpress.XtraEditors.PanelControl panelControl4;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraEditors.PanelControl panelControl5;
        private DevExpress.XtraEditors.SimpleButton btnSaveIo;
    }
}