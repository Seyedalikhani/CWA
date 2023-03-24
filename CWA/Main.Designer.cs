namespace CWA
{
    partial class Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.worstCellReportsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dashboardsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mAPToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.kPIZeroToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cRToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.availabilityToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.coreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.customerComplainToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.worstCellReportsToolStripMenuItem,
            this.dashboardsToolStripMenuItem,
            this.mAPToolStripMenuItem,
            this.kPIZeroToolStripMenuItem,
            this.cRToolStripMenuItem,
            this.exportToolStripMenuItem,
            this.availabilityToolStripMenuItem1,
            this.coreToolStripMenuItem,
            this.customerComplainToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1276, 24);
            this.menuStrip1.TabIndex = 21;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // worstCellReportsToolStripMenuItem
            // 
            this.worstCellReportsToolStripMenuItem.Name = "worstCellReportsToolStripMenuItem";
            this.worstCellReportsToolStripMenuItem.Size = new System.Drawing.Size(116, 20);
            this.worstCellReportsToolStripMenuItem.Text = "Worst Cell Reports";
            this.worstCellReportsToolStripMenuItem.Click += new System.EventHandler(this.worstCellReportsToolStripMenuItem_Click);
            // 
            // dashboardsToolStripMenuItem
            // 
            this.dashboardsToolStripMenuItem.Name = "dashboardsToolStripMenuItem";
            this.dashboardsToolStripMenuItem.Size = new System.Drawing.Size(81, 20);
            this.dashboardsToolStripMenuItem.Text = "Dashboards";
            this.dashboardsToolStripMenuItem.Click += new System.EventHandler(this.dashboardsToolStripMenuItem_Click);
            // 
            // mAPToolStripMenuItem
            // 
            this.mAPToolStripMenuItem.Name = "mAPToolStripMenuItem";
            this.mAPToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.mAPToolStripMenuItem.Text = "MAP";
            this.mAPToolStripMenuItem.Click += new System.EventHandler(this.mAPToolStripMenuItem_Click);
            // 
            // kPIZeroToolStripMenuItem
            // 
            this.kPIZeroToolStripMenuItem.Name = "kPIZeroToolStripMenuItem";
            this.kPIZeroToolStripMenuItem.Size = new System.Drawing.Size(63, 20);
            this.kPIZeroToolStripMenuItem.Text = "KPI Zero";
            this.kPIZeroToolStripMenuItem.Click += new System.EventHandler(this.kPIZeroToolStripMenuItem_Click);
            // 
            // cRToolStripMenuItem
            // 
            this.cRToolStripMenuItem.Name = "cRToolStripMenuItem";
            this.cRToolStripMenuItem.Size = new System.Drawing.Size(34, 20);
            this.cRToolStripMenuItem.Text = "CR";
            this.cRToolStripMenuItem.Click += new System.EventHandler(this.cRToolStripMenuItem_Click);
            // 
            // exportToolStripMenuItem
            // 
            this.exportToolStripMenuItem.Name = "exportToolStripMenuItem";
            this.exportToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.exportToolStripMenuItem.Text = "LTE";
            this.exportToolStripMenuItem.Click += new System.EventHandler(this.exportToolStripMenuItem_Click);
            // 
            // availabilityToolStripMenuItem1
            // 
            this.availabilityToolStripMenuItem1.Name = "availabilityToolStripMenuItem1";
            this.availabilityToolStripMenuItem1.Size = new System.Drawing.Size(77, 20);
            this.availabilityToolStripMenuItem1.Text = "Availability";
            this.availabilityToolStripMenuItem1.Click += new System.EventHandler(this.availabilityToolStripMenuItem1_Click);
            // 
            // coreToolStripMenuItem
            // 
            this.coreToolStripMenuItem.Name = "coreToolStripMenuItem";
            this.coreToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.coreToolStripMenuItem.Text = "Core";
            this.coreToolStripMenuItem.Click += new System.EventHandler(this.coreToolStripMenuItem_Click);
            // 
            // customerComplainToolStripMenuItem
            // 
            this.customerComplainToolStripMenuItem.Name = "customerComplainToolStripMenuItem";
            this.customerComplainToolStripMenuItem.Size = new System.Drawing.Size(130, 20);
            this.customerComplainToolStripMenuItem.Text = "Customer Complaint";
            this.customerComplainToolStripMenuItem.Click += new System.EventHandler(this.customerComplainToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(1276, 719);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NPM (Network Performance Monitoring)";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem worstCellReportsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dashboardsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mAPToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem kPIZeroToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cRToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem availabilityToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem coreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem customerComplainToolStripMenuItem;
    }
}

