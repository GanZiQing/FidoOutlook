namespace OutlookAutomation
{
    partial class HdbExport
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
            this.linkJsonFile = new System.Windows.Forms.Button();
            this.listView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.addNewProject = new System.Windows.Forms.Button();
            this.saveJson = new System.Windows.Forms.Button();
            this.importJson = new System.Windows.Forms.Button();
            this.editProjectButton = new System.Windows.Forms.Button();
            this.deleteProjectButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dispLinkedPath = new System.Windows.Forms.TextBox();
            this.unlinkJson = new System.Windows.Forms.Button();
            this.linkJsonGroup = new System.Windows.Forms.GroupBox();
            this.lastSavedLabel = new System.Windows.Forms.Label();
            this.exportEmailGroup = new System.Windows.Forms.GroupBox();
            this.exportSelectedOnly = new System.Windows.Forms.Button();
            this.exportAllProjects = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.hdbExportTabPage = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.linkJsonGroup.SuspendLayout();
            this.exportEmailGroup.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.hdbExportTabPage.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // linkJsonFile
            // 
            this.linkJsonFile.ForeColor = System.Drawing.Color.Black;
            this.linkJsonFile.Location = new System.Drawing.Point(9, 31);
            this.linkJsonFile.Margin = new System.Windows.Forms.Padding(6);
            this.linkJsonFile.Name = "linkJsonFile";
            this.linkJsonFile.Size = new System.Drawing.Size(238, 48);
            this.linkJsonFile.TabIndex = 1;
            this.linkJsonFile.Text = "Link Json File";
            this.linkJsonFile.UseVisualStyleBackColor = true;
            this.linkJsonFile.Click += new System.EventHandler(this.linkJsonFile_Click);
            // 
            // listView
            // 
            this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listView.ForeColor = System.Drawing.Color.Black;
            this.listView.HideSelection = false;
            this.listView.Location = new System.Drawing.Point(11, 35);
            this.listView.Margin = new System.Windows.Forms.Padding(6);
            this.listView.MultiSelect = false;
            this.listView.Name = "listView";
            this.listView.Size = new System.Drawing.Size(479, 299);
            this.listView.TabIndex = 1;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Project Name";
            this.columnHeader1.Width = 259;
            // 
            // addNewProject
            // 
            this.addNewProject.ForeColor = System.Drawing.Color.Black;
            this.addNewProject.Location = new System.Drawing.Point(11, 349);
            this.addNewProject.Margin = new System.Windows.Forms.Padding(6);
            this.addNewProject.Name = "addNewProject";
            this.addNewProject.Size = new System.Drawing.Size(238, 48);
            this.addNewProject.TabIndex = 2;
            this.addNewProject.Text = "Add New Project";
            this.addNewProject.UseVisualStyleBackColor = true;
            this.addNewProject.Click += new System.EventHandler(this.addNewProject_Click);
            // 
            // saveJson
            // 
            this.saveJson.ForeColor = System.Drawing.Color.Black;
            this.saveJson.Location = new System.Drawing.Point(257, 31);
            this.saveJson.Margin = new System.Windows.Forms.Padding(6);
            this.saveJson.Name = "saveJson";
            this.saveJson.Size = new System.Drawing.Size(238, 48);
            this.saveJson.TabIndex = 2;
            this.saveJson.Text = "Export Json File";
            this.saveJson.UseVisualStyleBackColor = true;
            this.saveJson.Click += new System.EventHandler(this.exportJson_Click);
            // 
            // importJson
            // 
            this.importJson.ForeColor = System.Drawing.Color.Black;
            this.importJson.Location = new System.Drawing.Point(9, 31);
            this.importJson.Margin = new System.Windows.Forms.Padding(6);
            this.importJson.Name = "importJson";
            this.importJson.Size = new System.Drawing.Size(238, 48);
            this.importJson.TabIndex = 1;
            this.importJson.Text = "Import Json File";
            this.importJson.UseVisualStyleBackColor = true;
            this.importJson.Click += new System.EventHandler(this.importJson_Click);
            // 
            // editProjectButton
            // 
            this.editProjectButton.ForeColor = System.Drawing.Color.Black;
            this.editProjectButton.Location = new System.Drawing.Point(13, 408);
            this.editProjectButton.Margin = new System.Windows.Forms.Padding(6);
            this.editProjectButton.Name = "editProjectButton";
            this.editProjectButton.Size = new System.Drawing.Size(238, 48);
            this.editProjectButton.TabIndex = 4;
            this.editProjectButton.Text = "Edit Project";
            this.editProjectButton.UseVisualStyleBackColor = true;
            this.editProjectButton.Click += new System.EventHandler(this.editProjectButton_Click);
            // 
            // deleteProjectButton
            // 
            this.deleteProjectButton.ForeColor = System.Drawing.Color.Black;
            this.deleteProjectButton.Location = new System.Drawing.Point(258, 349);
            this.deleteProjectButton.Margin = new System.Windows.Forms.Padding(6);
            this.deleteProjectButton.Name = "deleteProjectButton";
            this.deleteProjectButton.Size = new System.Drawing.Size(238, 48);
            this.deleteProjectButton.TabIndex = 3;
            this.deleteProjectButton.Text = "Delete Project";
            this.deleteProjectButton.UseVisualStyleBackColor = true;
            this.deleteProjectButton.Click += new System.EventHandler(this.deleteProjectButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.importJson);
            this.groupBox1.Controls.Add(this.saveJson);
            this.groupBox1.Location = new System.Drawing.Point(11, 685);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(502, 92);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Import/Export";
            // 
            // dispLinkedPath
            // 
            this.dispLinkedPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dispLinkedPath.BackColor = System.Drawing.SystemColors.Window;
            this.dispLinkedPath.ForeColor = System.Drawing.Color.Black;
            this.dispLinkedPath.Location = new System.Drawing.Point(9, 81);
            this.dispLinkedPath.Margin = new System.Windows.Forms.Padding(4);
            this.dispLinkedPath.Multiline = true;
            this.dispLinkedPath.Name = "dispLinkedPath";
            this.dispLinkedPath.ReadOnly = true;
            this.dispLinkedPath.Size = new System.Drawing.Size(482, 63);
            this.dispLinkedPath.TabIndex = 3;
            this.dispLinkedPath.TabStop = false;
            // 
            // unlinkJson
            // 
            this.unlinkJson.ForeColor = System.Drawing.Color.Black;
            this.unlinkJson.Location = new System.Drawing.Point(257, 31);
            this.unlinkJson.Margin = new System.Windows.Forms.Padding(6);
            this.unlinkJson.Name = "unlinkJson";
            this.unlinkJson.Size = new System.Drawing.Size(238, 48);
            this.unlinkJson.TabIndex = 2;
            this.unlinkJson.Text = "Unlink Json File";
            this.unlinkJson.UseVisualStyleBackColor = true;
            this.unlinkJson.Click += new System.EventHandler(this.unlinkJson_Click);
            // 
            // linkJsonGroup
            // 
            this.linkJsonGroup.Controls.Add(this.lastSavedLabel);
            this.linkJsonGroup.Controls.Add(this.linkJsonFile);
            this.linkJsonGroup.Controls.Add(this.unlinkJson);
            this.linkJsonGroup.Controls.Add(this.dispLinkedPath);
            this.linkJsonGroup.Location = new System.Drawing.Point(9, 489);
            this.linkJsonGroup.Margin = new System.Windows.Forms.Padding(4);
            this.linkJsonGroup.Name = "linkJsonGroup";
            this.linkJsonGroup.Padding = new System.Windows.Forms.Padding(4);
            this.linkJsonGroup.Size = new System.Drawing.Size(504, 188);
            this.linkJsonGroup.TabIndex = 5;
            this.linkJsonGroup.TabStop = false;
            this.linkJsonGroup.Text = "Link Json File";
            // 
            // lastSavedLabel
            // 
            this.lastSavedLabel.ForeColor = System.Drawing.Color.Black;
            this.lastSavedLabel.Location = new System.Drawing.Point(13, 151);
            this.lastSavedLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lastSavedLabel.Name = "lastSavedLabel";
            this.lastSavedLabel.Size = new System.Drawing.Size(482, 31);
            this.lastSavedLabel.TabIndex = 4;
            this.lastSavedLabel.Text = "Last saved: Never";
            // 
            // exportEmailGroup
            // 
            this.exportEmailGroup.Controls.Add(this.exportSelectedOnly);
            this.exportEmailGroup.Controls.Add(this.exportAllProjects);
            this.exportEmailGroup.Location = new System.Drawing.Point(9, 785);
            this.exportEmailGroup.Margin = new System.Windows.Forms.Padding(4);
            this.exportEmailGroup.Name = "exportEmailGroup";
            this.exportEmailGroup.Padding = new System.Windows.Forms.Padding(4);
            this.exportEmailGroup.Size = new System.Drawing.Size(502, 94);
            this.exportEmailGroup.TabIndex = 7;
            this.exportEmailGroup.TabStop = false;
            this.exportEmailGroup.Text = "Export Email";
            // 
            // exportSelectedOnly
            // 
            this.exportSelectedOnly.ForeColor = System.Drawing.Color.Black;
            this.exportSelectedOnly.Location = new System.Drawing.Point(9, 31);
            this.exportSelectedOnly.Margin = new System.Windows.Forms.Padding(6);
            this.exportSelectedOnly.Name = "exportSelectedOnly";
            this.exportSelectedOnly.Size = new System.Drawing.Size(238, 48);
            this.exportSelectedOnly.TabIndex = 1;
            this.exportSelectedOnly.Text = "Export Selected Project";
            this.exportSelectedOnly.UseVisualStyleBackColor = true;
            this.exportSelectedOnly.Click += new System.EventHandler(this.exportSelectedOnly_Click);
            // 
            // exportAllProjects
            // 
            this.exportAllProjects.ForeColor = System.Drawing.Color.Black;
            this.exportAllProjects.Location = new System.Drawing.Point(257, 31);
            this.exportAllProjects.Margin = new System.Windows.Forms.Padding(6);
            this.exportAllProjects.Name = "exportAllProjects";
            this.exportAllProjects.Size = new System.Drawing.Size(238, 48);
            this.exportAllProjects.TabIndex = 2;
            this.exportAllProjects.Text = "Export All Projects";
            this.exportAllProjects.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.hdbExportTabPage);
            this.tabControl1.Location = new System.Drawing.Point(6, 6);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(539, 1097);
            this.tabControl1.TabIndex = 15;
            // 
            // hdbExportTabPage
            // 
            this.hdbExportTabPage.Controls.Add(this.groupBox2);
            this.hdbExportTabPage.Controls.Add(this.exportEmailGroup);
            this.hdbExportTabPage.Controls.Add(this.linkJsonGroup);
            this.hdbExportTabPage.Controls.Add(this.groupBox1);
            this.hdbExportTabPage.Location = new System.Drawing.Point(4, 33);
            this.hdbExportTabPage.Margin = new System.Windows.Forms.Padding(4);
            this.hdbExportTabPage.Name = "hdbExportTabPage";
            this.hdbExportTabPage.Padding = new System.Windows.Forms.Padding(4);
            this.hdbExportTabPage.Size = new System.Drawing.Size(531, 1060);
            this.hdbExportTabPage.TabIndex = 0;
            this.hdbExportTabPage.Text = "Advance Export";
            this.hdbExportTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.listView);
            this.groupBox2.Controls.Add(this.addNewProject);
            this.groupBox2.Controls.Add(this.deleteProjectButton);
            this.groupBox2.Controls.Add(this.editProjectButton);
            this.groupBox2.Location = new System.Drawing.Point(9, 9);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(6);
            this.groupBox2.Size = new System.Drawing.Size(504, 471);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Define Projects";
            // 
            // HdbExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControl1);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "HdbExport";
            this.Size = new System.Drawing.Size(550, 1108);
            this.groupBox1.ResumeLayout(false);
            this.linkJsonGroup.ResumeLayout(false);
            this.linkJsonGroup.PerformLayout();
            this.exportEmailGroup.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.hdbExportTabPage.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button linkJsonFile;
        private System.Windows.Forms.ListView listView;
        private System.Windows.Forms.Button addNewProject;
        private System.Windows.Forms.Button saveJson;
        private System.Windows.Forms.Button importJson;
        private System.Windows.Forms.Button editProjectButton;
        private System.Windows.Forms.Button deleteProjectButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox dispLinkedPath;
        private System.Windows.Forms.Button unlinkJson;
        private System.Windows.Forms.GroupBox linkJsonGroup;
        private System.Windows.Forms.GroupBox exportEmailGroup;
        private System.Windows.Forms.Button exportSelectedOnly;
        private System.Windows.Forms.Button exportAllProjects;
        private System.Windows.Forms.Label lastSavedLabel;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage hdbExportTabPage;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ColumnHeader columnHeader1;
    }
}