
namespace RemarkableSync.OnenoteAddin
{
    partial class RmArchiveForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RmArchiveForm));
            this.tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.lblInfo = new System.Windows.Forms.Label();
            this.rmTreeView = new System.Windows.Forms.TreeView();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.chkArchiveAll = new System.Windows.Forms.CheckBox();
            this.tableLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel
            // 
            resources.ApplyResources(this.tableLayoutPanel, "tableLayoutPanel");
            this.tableLayoutPanel.Controls.Add(this.lblInfo, 0, 1);
            this.tableLayoutPanel.Controls.Add(this.rmTreeView, 0, 0);
            this.tableLayoutPanel.Controls.Add(this.btnOk, 0, 3);
            this.tableLayoutPanel.Controls.Add(this.btnCancel, 1, 3);
            this.tableLayoutPanel.Controls.Add(this.chkArchiveAll, 0, 2);
            this.tableLayoutPanel.Name = "tableLayoutPanel";
            // 
            // lblInfo
            // 
            resources.ApplyResources(this.lblInfo, "lblInfo");
            this.tableLayoutPanel.SetColumnSpan(this.lblInfo, 2);
            this.lblInfo.Name = "lblInfo";
            // 
            // rmTreeView
            // 
            this.tableLayoutPanel.SetColumnSpan(this.rmTreeView, 2);
            resources.ApplyResources(this.rmTreeView, "rmTreeView");
            this.rmTreeView.Name = "rmTreeView";
            this.rmTreeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.rmTreeView_AfterSelect);
            // 
            // btnOk
            // 
            resources.ApplyResources(this.btnOk, "btnOk");
            this.btnOk.Name = "btnOk";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // chkArchiveAll
            // 
            resources.ApplyResources(this.chkArchiveAll, "chkArchiveAll");
            this.chkArchiveAll.Name = "chkArchiveAll";
            this.chkArchiveAll.UseVisualStyleBackColor = true;
            this.chkArchiveAll.CheckedChanged += new System.EventHandler(this.chkArchiveAll_CheckedChanged);
            // 
            // RmArchiveForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel);
            this.Name = "RmArchiveForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.RmDownloadForm_FormClosing);
            this.tableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.TreeView rmTreeView;
        private System.Windows.Forms.CheckBox chkArchiveAll;
    }
}