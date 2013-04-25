namespace Gardiner.LoadedProjects
{
    partial class frmProgress
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCurrent = new System.Windows.Forms.Label();
            this.lblAction = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 93);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(558, 23);
            this.progressBar1.TabIndex = 0;
            this.progressBar1.UseWaitCursor = true;
            // 
            // lblCurrent
            // 
            this.lblCurrent.Location = new System.Drawing.Point(12, 51);
            this.lblCurrent.Name = "lblCurrent";
            this.lblCurrent.Size = new System.Drawing.Size(558, 39);
            this.lblCurrent.TabIndex = 1;
            this.lblCurrent.UseWaitCursor = true;
            // 
            // lblAction
            // 
            this.lblAction.Location = new System.Drawing.Point(12, 13);
            this.lblAction.Name = "lblAction";
            this.lblAction.Size = new System.Drawing.Size(558, 26);
            this.lblAction.TabIndex = 2;
            this.lblAction.UseWaitCursor = true;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // frmProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(582, 131);
            this.ControlBox = false;
            this.Controls.Add(this.lblAction);
            this.Controls.Add(this.lblCurrent);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmProgress";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Loaded Projects";
            this.TopMost = true;
            this.UseWaitCursor = true;
            this.Load += new System.EventHandler(this.frmProgress_Load);
            this.Shown += new System.EventHandler(this.frmProgress_Shown);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCurrent;
        private System.Windows.Forms.Label lblAction;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}