namespace OutlookAutomation
{
    partial class PrintPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.exportMSG = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.exportAllSelected = new System.Windows.Forms.Button();
            this.setFolder = new System.Windows.Forms.Button();
            this.openFolder = new System.Windows.Forms.Button();
            this.dispBaseFolder = new System.Windows.Forms.TextBox();
            this.exportOptionsGroupBox = new System.Windows.Forms.GroupBox();
            this.moveItemsCheck = new System.Windows.Forms.CheckBox();
            this.breakOnErrorCheck = new System.Windows.Forms.CheckBox();
            this.saveSettings = new System.Windows.Forms.Button();
            this.dispMaxSubjectLength = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.shortenSubjectCheck = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.skipEmbeddedCheck = new System.Windows.Forms.CheckBox();
            this.exportWordCheck = new System.Windows.Forms.CheckBox();
            this.exportHtmlCheck = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.exportMsgCheck = new System.Windows.Forms.CheckBox();
            this.exportPdfCheck = new System.Windows.Forms.CheckBox();
            this.dateFolderCheck = new System.Windows.Forms.CheckBox();
            this.subjectFolderCheck = new System.Windows.Forms.CheckBox();
            this.exportAttachmentsCheck = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tabPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.exportOptionsGroupBox.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.exportOptionsGroupBox);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(6);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(6);
            this.tabPage1.Size = new System.Drawing.Size(531, 1798);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Export";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.exportMSG);
            this.groupBox2.Location = new System.Drawing.Point(11, 810);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(502, 133);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Export By File";
            // 
            // exportMSG
            // 
            this.exportMSG.ForeColor = System.Drawing.Color.Black;
            this.exportMSG.Location = new System.Drawing.Point(141, 31);
            this.exportMSG.Margin = new System.Windows.Forms.Padding(6);
            this.exportMSG.Name = "exportMSG";
            this.exportMSG.Size = new System.Drawing.Size(220, 48);
            this.exportMSG.TabIndex = 14;
            this.exportMSG.Text = "Export .msg";
            this.exportMSG.UseVisualStyleBackColor = true;
            this.exportMSG.Click += new System.EventHandler(this.exportMSG_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.exportAllSelected);
            this.groupBox1.Controls.Add(this.setFolder);
            this.groupBox1.Controls.Add(this.openFolder);
            this.groupBox1.Controls.Add(this.dispBaseFolder);
            this.groupBox1.Location = new System.Drawing.Point(11, 594);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox1.Size = new System.Drawing.Size(502, 207);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Export Operations";
            // 
            // exportAllSelected
            // 
            this.exportAllSelected.ForeColor = System.Drawing.Color.Black;
            this.exportAllSelected.Location = new System.Drawing.Point(141, 142);
            this.exportAllSelected.Margin = new System.Windows.Forms.Padding(6);
            this.exportAllSelected.Name = "exportAllSelected";
            this.exportAllSelected.Size = new System.Drawing.Size(220, 48);
            this.exportAllSelected.TabIndex = 4;
            this.exportAllSelected.Text = "Export All Selected";
            this.exportAllSelected.UseVisualStyleBackColor = true;
            this.exportAllSelected.Click += new System.EventHandler(this.exportAllSelected_Click);
            // 
            // setFolder
            // 
            this.setFolder.ForeColor = System.Drawing.Color.Black;
            this.setFolder.Location = new System.Drawing.Point(11, 35);
            this.setFolder.Margin = new System.Windows.Forms.Padding(6);
            this.setFolder.Name = "setFolder";
            this.setFolder.Size = new System.Drawing.Size(220, 48);
            this.setFolder.TabIndex = 1;
            this.setFolder.Text = "Set Folder";
            this.setFolder.UseVisualStyleBackColor = true;
            this.setFolder.Click += new System.EventHandler(this.setFolder_Click);
            // 
            // openFolder
            // 
            this.openFolder.ForeColor = System.Drawing.Color.Black;
            this.openFolder.Location = new System.Drawing.Point(271, 35);
            this.openFolder.Margin = new System.Windows.Forms.Padding(6);
            this.openFolder.Name = "openFolder";
            this.openFolder.Size = new System.Drawing.Size(220, 48);
            this.openFolder.TabIndex = 2;
            this.openFolder.Text = "Open Folder";
            this.openFolder.UseVisualStyleBackColor = true;
            this.openFolder.Click += new System.EventHandler(this.openFolder_Click);
            // 
            // dispBaseFolder
            // 
            this.dispBaseFolder.Location = new System.Drawing.Point(11, 94);
            this.dispBaseFolder.Margin = new System.Windows.Forms.Padding(6);
            this.dispBaseFolder.Name = "dispBaseFolder";
            this.dispBaseFolder.Size = new System.Drawing.Size(477, 29);
            this.dispBaseFolder.TabIndex = 3;
            // 
            // exportOptionsGroupBox
            // 
            this.exportOptionsGroupBox.Controls.Add(this.moveItemsCheck);
            this.exportOptionsGroupBox.Controls.Add(this.breakOnErrorCheck);
            this.exportOptionsGroupBox.Controls.Add(this.saveSettings);
            this.exportOptionsGroupBox.Controls.Add(this.dispMaxSubjectLength);
            this.exportOptionsGroupBox.Controls.Add(this.label4);
            this.exportOptionsGroupBox.Controls.Add(this.shortenSubjectCheck);
            this.exportOptionsGroupBox.Controls.Add(this.label3);
            this.exportOptionsGroupBox.Controls.Add(this.skipEmbeddedCheck);
            this.exportOptionsGroupBox.Controls.Add(this.exportWordCheck);
            this.exportOptionsGroupBox.Controls.Add(this.exportHtmlCheck);
            this.exportOptionsGroupBox.Controls.Add(this.label2);
            this.exportOptionsGroupBox.Controls.Add(this.label1);
            this.exportOptionsGroupBox.Controls.Add(this.exportMsgCheck);
            this.exportOptionsGroupBox.Controls.Add(this.exportPdfCheck);
            this.exportOptionsGroupBox.Controls.Add(this.dateFolderCheck);
            this.exportOptionsGroupBox.Controls.Add(this.subjectFolderCheck);
            this.exportOptionsGroupBox.Controls.Add(this.exportAttachmentsCheck);
            this.exportOptionsGroupBox.Location = new System.Drawing.Point(11, 11);
            this.exportOptionsGroupBox.Margin = new System.Windows.Forms.Padding(6);
            this.exportOptionsGroupBox.Name = "exportOptionsGroupBox";
            this.exportOptionsGroupBox.Padding = new System.Windows.Forms.Padding(6);
            this.exportOptionsGroupBox.Size = new System.Drawing.Size(502, 572);
            this.exportOptionsGroupBox.TabIndex = 1;
            this.exportOptionsGroupBox.TabStop = false;
            this.exportOptionsGroupBox.Text = "Export Options";
            // 
            // moveItemsCheck
            // 
            this.moveItemsCheck.Checked = true;
            this.moveItemsCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.moveItemsCheck.ForeColor = System.Drawing.Color.Black;
            this.moveItemsCheck.Location = new System.Drawing.Point(11, 463);
            this.moveItemsCheck.Margin = new System.Windows.Forms.Padding(4);
            this.moveItemsCheck.Name = "moveItemsCheck";
            this.moveItemsCheck.Size = new System.Drawing.Size(480, 31);
            this.moveItemsCheck.TabIndex = 12;
            this.moveItemsCheck.Text = "Move Email To Export Folders";
            this.moveItemsCheck.UseVisualStyleBackColor = true;
            // 
            // breakOnErrorCheck
            // 
            this.breakOnErrorCheck.Checked = true;
            this.breakOnErrorCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.breakOnErrorCheck.ForeColor = System.Drawing.Color.Black;
            this.breakOnErrorCheck.Location = new System.Drawing.Point(11, 425);
            this.breakOnErrorCheck.Margin = new System.Windows.Forms.Padding(6);
            this.breakOnErrorCheck.Name = "breakOnErrorCheck";
            this.breakOnErrorCheck.Size = new System.Drawing.Size(480, 31);
            this.breakOnErrorCheck.TabIndex = 11;
            this.breakOnErrorCheck.Text = "Break On Error";
            this.breakOnErrorCheck.UseVisualStyleBackColor = true;
            // 
            // saveSettings
            // 
            this.saveSettings.ForeColor = System.Drawing.Color.Black;
            this.saveSettings.Location = new System.Drawing.Point(141, 504);
            this.saveSettings.Margin = new System.Windows.Forms.Padding(6);
            this.saveSettings.Name = "saveSettings";
            this.saveSettings.Size = new System.Drawing.Size(220, 48);
            this.saveSettings.TabIndex = 13;
            this.saveSettings.Text = "Save Settings";
            this.saveSettings.UseVisualStyleBackColor = true;
            this.saveSettings.Click += new System.EventHandler(this.saveSettingsButton_Click);
            // 
            // dispMaxSubjectLength
            // 
            this.dispMaxSubjectLength.Location = new System.Drawing.Point(422, 377);
            this.dispMaxSubjectLength.Margin = new System.Windows.Forms.Padding(6);
            this.dispMaxSubjectLength.Name = "dispMaxSubjectLength";
            this.dispMaxSubjectLength.Size = new System.Drawing.Size(48, 29);
            this.dispMaxSubjectLength.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(218, 377);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(192, 37);
            this.label4.TabIndex = 13;
            this.label4.Text = "Max Subject Length:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // shortenSubjectCheck
            // 
            this.shortenSubjectCheck.ForeColor = System.Drawing.Color.Black;
            this.shortenSubjectCheck.Location = new System.Drawing.Point(11, 382);
            this.shortenSubjectCheck.Margin = new System.Windows.Forms.Padding(6);
            this.shortenSubjectCheck.Name = "shortenSubjectCheck";
            this.shortenSubjectCheck.Size = new System.Drawing.Size(196, 31);
            this.shortenSubjectCheck.TabIndex = 9;
            this.shortenSubjectCheck.Text = "Shorten Subject";
            this.shortenSubjectCheck.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(11, 297);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(480, 37);
            this.label3.TabIndex = 11;
            this.label3.Text = "Other Options:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // skipEmbeddedCheck
            // 
            this.skipEmbeddedCheck.ForeColor = System.Drawing.Color.Black;
            this.skipEmbeddedCheck.Location = new System.Drawing.Point(11, 340);
            this.skipEmbeddedCheck.Margin = new System.Windows.Forms.Padding(6);
            this.skipEmbeddedCheck.Name = "skipEmbeddedCheck";
            this.skipEmbeddedCheck.Size = new System.Drawing.Size(480, 31);
            this.skipEmbeddedCheck.TabIndex = 8;
            this.skipEmbeddedCheck.Text = "Skip Embeded Attachments";
            this.skipEmbeddedCheck.UseVisualStyleBackColor = true;
            // 
            // exportWordCheck
            // 
            this.exportWordCheck.ForeColor = System.Drawing.Color.Black;
            this.exportWordCheck.Location = new System.Drawing.Point(16, 199);
            this.exportWordCheck.Margin = new System.Windows.Forms.Padding(6);
            this.exportWordCheck.Name = "exportWordCheck";
            this.exportWordCheck.Size = new System.Drawing.Size(211, 31);
            this.exportWordCheck.TabIndex = 4;
            this.exportWordCheck.Text = ".docx";
            this.exportWordCheck.UseVisualStyleBackColor = true;
            // 
            // exportHtmlCheck
            // 
            this.exportHtmlCheck.ForeColor = System.Drawing.Color.Black;
            this.exportHtmlCheck.Location = new System.Drawing.Point(16, 114);
            this.exportHtmlCheck.Margin = new System.Windows.Forms.Padding(6);
            this.exportHtmlCheck.Name = "exportHtmlCheck";
            this.exportHtmlCheck.Size = new System.Drawing.Size(211, 31);
            this.exportHtmlCheck.TabIndex = 2;
            this.exportHtmlCheck.Text = ".html";
            this.exportHtmlCheck.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(233, 30);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(132, 37);
            this.label2.TabIndex = 6;
            this.label2.Text = "Folders:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label1
            // 
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(11, 30);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 37);
            this.label1.TabIndex = 5;
            this.label1.Text = "Export Types:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // exportMsgCheck
            // 
            this.exportMsgCheck.Checked = true;
            this.exportMsgCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.exportMsgCheck.ForeColor = System.Drawing.Color.Black;
            this.exportMsgCheck.Location = new System.Drawing.Point(16, 157);
            this.exportMsgCheck.Margin = new System.Windows.Forms.Padding(6);
            this.exportMsgCheck.Name = "exportMsgCheck";
            this.exportMsgCheck.Size = new System.Drawing.Size(211, 31);
            this.exportMsgCheck.TabIndex = 3;
            this.exportMsgCheck.Text = ".msg";
            this.exportMsgCheck.UseVisualStyleBackColor = true;
            // 
            // exportPdfCheck
            // 
            this.exportPdfCheck.ForeColor = System.Drawing.Color.Black;
            this.exportPdfCheck.Location = new System.Drawing.Point(16, 242);
            this.exportPdfCheck.Margin = new System.Windows.Forms.Padding(6);
            this.exportPdfCheck.Name = "exportPdfCheck";
            this.exportPdfCheck.Size = new System.Drawing.Size(211, 31);
            this.exportPdfCheck.TabIndex = 5;
            this.exportPdfCheck.Text = ".pdf";
            this.exportPdfCheck.UseVisualStyleBackColor = true;
            // 
            // dateFolderCheck
            // 
            this.dateFolderCheck.Checked = true;
            this.dateFolderCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.dateFolderCheck.ForeColor = System.Drawing.Color.Black;
            this.dateFolderCheck.Location = new System.Drawing.Point(238, 114);
            this.dateFolderCheck.Margin = new System.Windows.Forms.Padding(6);
            this.dateFolderCheck.Name = "dateFolderCheck";
            this.dateFolderCheck.Size = new System.Drawing.Size(211, 31);
            this.dateFolderCheck.TabIndex = 7;
            this.dateFolderCheck.Text = "Date Time Folder";
            this.dateFolderCheck.UseVisualStyleBackColor = true;
            // 
            // subjectFolderCheck
            // 
            this.subjectFolderCheck.Checked = true;
            this.subjectFolderCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.subjectFolderCheck.ForeColor = System.Drawing.Color.Black;
            this.subjectFolderCheck.Location = new System.Drawing.Point(238, 72);
            this.subjectFolderCheck.Margin = new System.Windows.Forms.Padding(6);
            this.subjectFolderCheck.Name = "subjectFolderCheck";
            this.subjectFolderCheck.Size = new System.Drawing.Size(211, 31);
            this.subjectFolderCheck.TabIndex = 6;
            this.subjectFolderCheck.Text = "Subject Folder";
            this.subjectFolderCheck.UseVisualStyleBackColor = true;
            // 
            // exportAttachmentsCheck
            // 
            this.exportAttachmentsCheck.Checked = true;
            this.exportAttachmentsCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.exportAttachmentsCheck.ForeColor = System.Drawing.Color.Black;
            this.exportAttachmentsCheck.Location = new System.Drawing.Point(16, 72);
            this.exportAttachmentsCheck.Margin = new System.Windows.Forms.Padding(6);
            this.exportAttachmentsCheck.Name = "exportAttachmentsCheck";
            this.exportAttachmentsCheck.Size = new System.Drawing.Size(211, 31);
            this.exportAttachmentsCheck.TabIndex = 1;
            this.exportAttachmentsCheck.Text = "Attachments";
            this.exportAttachmentsCheck.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(6, 6);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(6);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(539, 1835);
            this.tabControl1.TabIndex = 0;
            // 
            // PrintPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "PrintPane";
            this.Size = new System.Drawing.Size(550, 1846);
            this.tabPage1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.exportOptionsGroupBox.ResumeLayout(false);
            this.exportOptionsGroupBox.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.GroupBox exportOptionsGroupBox;
        private System.Windows.Forms.CheckBox exportPdfCheck;
        private System.Windows.Forms.CheckBox dateFolderCheck;
        private System.Windows.Forms.CheckBox subjectFolderCheck;
        private System.Windows.Forms.CheckBox exportAttachmentsCheck;
        private System.Windows.Forms.CheckBox exportMsgCheck;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox exportHtmlCheck;
        private System.Windows.Forms.CheckBox exportWordCheck;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button exportAllSelected;
        private System.Windows.Forms.CheckBox skipEmbeddedCheck;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox shortenSubjectCheck;
        private System.Windows.Forms.TextBox dispMaxSubjectLength;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button saveSettings;
        private System.Windows.Forms.Button openFolder;
        private System.Windows.Forms.Button setFolder;
        private System.Windows.Forms.TextBox dispBaseFolder;
        private System.Windows.Forms.CheckBox breakOnErrorCheck;
        private System.Windows.Forms.CheckBox moveItemsCheck;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button exportMSG;
    }
}
