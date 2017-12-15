namespace BurnMedia
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.devicesComboBox = new System.Windows.Forms.ComboBox();
            this.selectCdLabel = new System.Windows.Forms.Label();
            this.backgroundBurnWorker = new System.ComponentModel.BackgroundWorker();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.checkBoxEject = new System.Windows.Forms.CheckBox();
            this.textBoxLabel = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonRemoveFiles = new System.Windows.Forms.Button();
            this.buttonAddFolders = new System.Windows.Forms.Button();
            this.buttonAddFiles = new System.Windows.Forms.Button();
            this.listBoxFiles = new System.Windows.Forms.ListBox();
            this.labelMediaType = new System.Windows.Forms.Label();
            this.buttonDetectMedia = new System.Windows.Forms.Button();
            this.statusProgressBar = new System.Windows.Forms.ProgressBar();
            this.labelStatusText = new System.Windows.Forms.Label();
            this.buttonBurn = new System.Windows.Forms.Button();
            this.BurnerNotifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.FreeSpace = new System.Windows.Forms.Label();
            this.TotalSpace = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // devicesComboBox
            // 
            this.devicesComboBox.FormattingEnabled = true;
            this.devicesComboBox.Location = new System.Drawing.Point(76, 14);
            this.devicesComboBox.Name = "devicesComboBox";
            this.devicesComboBox.Size = new System.Drawing.Size(122, 21);
            this.devicesComboBox.TabIndex = 1;
            this.devicesComboBox.SelectedIndexChanged += new System.EventHandler(this.DevicesComboBox_SelectedIndexChanged);
            this.devicesComboBox.Format += new System.Windows.Forms.ListControlConvertEventHandler(this.DevicesComboBox_Format);
            // 
            // selectCdLabel
            // 
            this.selectCdLabel.AutoSize = true;
            this.selectCdLabel.Location = new System.Drawing.Point(12, 17);
            this.selectCdLabel.Name = "selectCdLabel";
            this.selectCdLabel.Size = new System.Drawing.Size(58, 13);
            this.selectCdLabel.TabIndex = 2;
            this.selectCdLabel.Text = "Select CD:";
            // 
            // backgroundBurnWorker
            // 
            this.backgroundBurnWorker.WorkerReportsProgress = true;
            this.backgroundBurnWorker.WorkerSupportsCancellation = true;
            this.backgroundBurnWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundBurnWorker_DoWork);
            this.backgroundBurnWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BackgroundBurnWorker_ProgressChanged);
            this.backgroundBurnWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundBurnWorker_RunWorkerCompleted);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "All Files (*.*)|*.*";
            // 
            // checkBoxEject
            // 
            this.checkBoxEject.AutoSize = true;
            this.checkBoxEject.Checked = true;
            this.checkBoxEject.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxEject.Location = new System.Drawing.Point(124, 217);
            this.checkBoxEject.Name = "checkBoxEject";
            this.checkBoxEject.Size = new System.Drawing.Size(74, 17);
            this.checkBoxEject.TabIndex = 22;
            this.checkBoxEject.Text = "Eject after";
            this.checkBoxEject.UseVisualStyleBackColor = true;
            // 
            // textBoxLabel
            // 
            this.textBoxLabel.Location = new System.Drawing.Point(76, 246);
            this.textBoxLabel.Name = "textBoxLabel";
            this.textBoxLabel.Size = new System.Drawing.Size(226, 20);
            this.textBoxLabel.TabIndex = 21;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 249);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 20;
            this.label2.Text = "Name disk:";
            // 
            // buttonRemoveFiles
            // 
            this.buttonRemoveFiles.Location = new System.Drawing.Point(204, 184);
            this.buttonRemoveFiles.Name = "buttonRemoveFiles";
            this.buttonRemoveFiles.Size = new System.Drawing.Size(98, 27);
            this.buttonRemoveFiles.TabIndex = 19;
            this.buttonRemoveFiles.Text = "Delete";
            this.buttonRemoveFiles.UseVisualStyleBackColor = true;
            this.buttonRemoveFiles.Click += new System.EventHandler(this.ButtonRemoveFiles_Click);
            // 
            // buttonAddFolders
            // 
            this.buttonAddFolders.Location = new System.Drawing.Point(109, 184);
            this.buttonAddFolders.Name = "buttonAddFolders";
            this.buttonAddFolders.Size = new System.Drawing.Size(89, 27);
            this.buttonAddFolders.TabIndex = 18;
            this.buttonAddFolders.Text = "Select Dir";
            this.buttonAddFolders.UseVisualStyleBackColor = true;
            this.buttonAddFolders.Click += new System.EventHandler(this.ButtonAddFolders_Click);
            // 
            // buttonAddFiles
            // 
            this.buttonAddFiles.Location = new System.Drawing.Point(12, 184);
            this.buttonAddFiles.Name = "buttonAddFiles";
            this.buttonAddFiles.Size = new System.Drawing.Size(91, 27);
            this.buttonAddFiles.TabIndex = 17;
            this.buttonAddFiles.Text = "Select Files";
            this.buttonAddFiles.UseVisualStyleBackColor = true;
            this.buttonAddFiles.Click += new System.EventHandler(this.ButtonAddFiles_Click);
            // 
            // listBoxFiles
            // 
            this.listBoxFiles.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.listBoxFiles.FormattingEnabled = true;
            this.listBoxFiles.ItemHeight = 24;
            this.listBoxFiles.Location = new System.Drawing.Point(15, 41);
            this.listBoxFiles.Name = "listBoxFiles";
            this.listBoxFiles.Size = new System.Drawing.Size(287, 124);
            this.listBoxFiles.TabIndex = 16;
            this.listBoxFiles.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ListBoxFiles_DrawItem);
            this.listBoxFiles.SelectedIndexChanged += new System.EventHandler(this.ListBoxFiles_SelectedIndexChanged);
            // 
            // labelMediaType
            // 
            this.labelMediaType.Location = new System.Drawing.Point(15, 214);
            this.labelMediaType.Name = "labelMediaType";
            this.labelMediaType.Size = new System.Drawing.Size(106, 26);
            this.labelMediaType.TabIndex = 24;
            // 
            // buttonDetectMedia
            // 
            this.buttonDetectMedia.Location = new System.Drawing.Point(204, 12);
            this.buttonDetectMedia.Name = "buttonDetectMedia";
            this.buttonDetectMedia.Size = new System.Drawing.Size(98, 27);
            this.buttonDetectMedia.TabIndex = 23;
            this.buttonDetectMedia.Text = "Detect Media";
            this.buttonDetectMedia.UseVisualStyleBackColor = true;
            this.buttonDetectMedia.Click += new System.EventHandler(this.ButtonDetectMedia_Click);
            // 
            // statusProgressBar
            // 
            this.statusProgressBar.Location = new System.Drawing.Point(12, 330);
            this.statusProgressBar.Name = "statusProgressBar";
            this.statusProgressBar.Size = new System.Drawing.Size(290, 30);
            this.statusProgressBar.TabIndex = 27;
            // 
            // labelStatusText
            // 
            this.labelStatusText.Location = new System.Drawing.Point(15, 307);
            this.labelStatusText.Name = "labelStatusText";
            this.labelStatusText.Size = new System.Drawing.Size(287, 20);
            this.labelStatusText.TabIndex = 26;
            this.labelStatusText.Text = "status";
            this.labelStatusText.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // buttonBurn
            // 
            this.buttonBurn.Location = new System.Drawing.Point(12, 276);
            this.buttonBurn.Name = "buttonBurn";
            this.buttonBurn.Size = new System.Drawing.Size(290, 28);
            this.buttonBurn.TabIndex = 25;
            this.buttonBurn.Text = "&Burn";
            this.buttonBurn.UseVisualStyleBackColor = true;
            this.buttonBurn.Click += new System.EventHandler(this.ButtonBurn_Click);
            // 
            // BurnerNotifyIcon
            // 
            this.BurnerNotifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.BurnerNotifyIcon.BalloonTipTitle = "Burner";
            this.BurnerNotifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("BurnerNotifyIcon.Icon")));
            this.BurnerNotifyIcon.Text = "Burner";
            this.BurnerNotifyIcon.Visible = true;
            this.BurnerNotifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.BurnerNotifyIcon_MouseDoubleClick);
            // 
            // FreeSpace
            // 
            this.FreeSpace.AutoSize = true;
            this.FreeSpace.Location = new System.Drawing.Point(167, 168);
            this.FreeSpace.Name = "FreeSpace";
            this.FreeSpace.Size = new System.Drawing.Size(0, 13);
            this.FreeSpace.TabIndex = 28;
            // 
            // TotalSpace
            // 
            this.TotalSpace.AutoSize = true;
            this.TotalSpace.Location = new System.Drawing.Point(15, 168);
            this.TotalSpace.Name = "TotalSpace";
            this.TotalSpace.Size = new System.Drawing.Size(0, 13);
            this.TotalSpace.TabIndex = 29;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(324, 372);
            this.Controls.Add(this.TotalSpace);
            this.Controls.Add(this.FreeSpace);
            this.Controls.Add(this.statusProgressBar);
            this.Controls.Add(this.labelStatusText);
            this.Controls.Add(this.buttonBurn);
            this.Controls.Add(this.labelMediaType);
            this.Controls.Add(this.buttonDetectMedia);
            this.Controls.Add(this.checkBoxEject);
            this.Controls.Add(this.textBoxLabel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonRemoveFiles);
            this.Controls.Add(this.buttonAddFolders);
            this.Controls.Add(this.buttonAddFiles);
            this.Controls.Add(this.listBoxFiles);
            this.Controls.Add(this.selectCdLabel);
            this.Controls.Add(this.devicesComboBox);
            this.Name = "MainForm";
            this.Text = "Burner";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Resize += new System.EventHandler(this.MainForm_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox devicesComboBox;
        private System.Windows.Forms.Label selectCdLabel;
        private System.ComponentModel.BackgroundWorker backgroundBurnWorker;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.CheckBox checkBoxEject;
        private System.Windows.Forms.TextBox textBoxLabel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonRemoveFiles;
        private System.Windows.Forms.Button buttonAddFolders;
        private System.Windows.Forms.Button buttonAddFiles;
        private System.Windows.Forms.ListBox listBoxFiles;
        private System.Windows.Forms.Label labelMediaType;
        private System.Windows.Forms.Button buttonDetectMedia;
        private System.Windows.Forms.ProgressBar statusProgressBar;
        private System.Windows.Forms.Label labelStatusText;
        private System.Windows.Forms.Button buttonBurn;
        private System.Windows.Forms.NotifyIcon BurnerNotifyIcon;
        private System.Windows.Forms.Label FreeSpace;
        private System.Windows.Forms.Label TotalSpace;
    }
}

