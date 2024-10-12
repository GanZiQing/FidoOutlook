namespace OutlookAutomation
{
    partial class CreateFilters
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
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.recipient = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sender = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.folderPath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.saveTable = new System.Windows.Forms.Button();
            this.loadTable = new System.Windows.Forms.Button();
            this.clearTable = new System.Windows.Forms.Button();
            this.checkTableButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.dispProjectName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.deleteRow = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.alwaysTopCheck = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.name,
            this.recipient,
            this.sender,
            this.folderPath});
            this.dataGridView.Location = new System.Drawing.Point(15, 55);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidth = 72;
            this.dataGridView.Size = new System.Drawing.Size(857, 327);
            this.dataGridView.TabIndex = 2;
            // 
            // name
            // 
            this.name.FillWeight = 25F;
            this.name.HeaderText = "Filter Name";
            this.name.MinimumWidth = 60;
            this.name.Name = "name";
            // 
            // recipient
            // 
            this.recipient.HeaderText = "Recipient";
            this.recipient.MinimumWidth = 9;
            this.recipient.Name = "recipient";
            // 
            // sender
            // 
            this.sender.HeaderText = "Sender";
            this.sender.MinimumWidth = 9;
            this.sender.Name = "sender";
            // 
            // folderPath
            // 
            this.folderPath.FillWeight = 200F;
            this.folderPath.HeaderText = "Folder Path";
            this.folderPath.MinimumWidth = 9;
            this.folderPath.Name = "folderPath";
            // 
            // saveTable
            // 
            this.saveTable.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.saveTable.Location = new System.Drawing.Point(301, 388);
            this.saveTable.Name = "saveTable";
            this.saveTable.Size = new System.Drawing.Size(100, 26);
            this.saveTable.TabIndex = 6;
            this.saveTable.Text = "Export Table";
            this.saveTable.UseVisualStyleBackColor = true;
            this.saveTable.Click += new System.EventHandler(this.saveTable_Click);
            // 
            // loadTable
            // 
            this.loadTable.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.loadTable.Location = new System.Drawing.Point(406, 388);
            this.loadTable.Name = "loadTable";
            this.loadTable.Size = new System.Drawing.Size(90, 26);
            this.loadTable.TabIndex = 7;
            this.loadTable.Text = "Load Table";
            this.loadTable.UseVisualStyleBackColor = true;
            this.loadTable.Click += new System.EventHandler(this.loadTable_Click);
            // 
            // clearTable
            // 
            this.clearTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.clearTable.Location = new System.Drawing.Point(204, 388);
            this.clearTable.Name = "clearTable";
            this.clearTable.Size = new System.Drawing.Size(90, 26);
            this.clearTable.TabIndex = 5;
            this.clearTable.Text = "ClearTable";
            this.clearTable.UseVisualStyleBackColor = true;
            this.clearTable.Click += new System.EventHandler(this.clearTable_Click);
            // 
            // checkTableButton
            // 
            this.checkTableButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkTableButton.Location = new System.Drawing.Point(12, 388);
            this.checkTableButton.Name = "checkTableButton";
            this.checkTableButton.Size = new System.Drawing.Size(90, 26);
            this.checkTableButton.TabIndex = 3;
            this.checkTableButton.Text = "Check Table";
            this.checkTableButton.UseVisualStyleBackColor = true;
            this.checkTableButton.Click += new System.EventHandler(this.checkTableButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Location = new System.Drawing.Point(782, 388);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(90, 26);
            this.cancelButton.TabIndex = 9;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // dispProjectName
            // 
            this.dispProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dispProjectName.Location = new System.Drawing.Point(15, 29);
            this.dispProjectName.Name = "dispProjectName";
            this.dispProjectName.Size = new System.Drawing.Size(857, 20);
            this.dispProjectName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(249, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "Project Name:";
            // 
            // deleteRow
            // 
            this.deleteRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.deleteRow.Location = new System.Drawing.Point(108, 388);
            this.deleteRow.Name = "deleteRow";
            this.deleteRow.Size = new System.Drawing.Size(90, 26);
            this.deleteRow.TabIndex = 4;
            this.deleteRow.Text = "Delete Row";
            this.deleteRow.UseVisualStyleBackColor = true;
            this.deleteRow.Click += new System.EventHandler(this.deleteRow_Click);
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.Location = new System.Drawing.Point(686, 388);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(90, 26);
            this.okButton.TabIndex = 8;
            this.okButton.Text = "Ok";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // alwaysTopCheck
            // 
            this.alwaysTopCheck.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.alwaysTopCheck.AutoSize = true;
            this.alwaysTopCheck.Location = new System.Drawing.Point(779, 4);
            this.alwaysTopCheck.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.alwaysTopCheck.Name = "alwaysTopCheck";
            this.alwaysTopCheck.Size = new System.Drawing.Size(98, 17);
            this.alwaysTopCheck.TabIndex = 10;
            this.alwaysTopCheck.Text = "Always On Top";
            this.alwaysTopCheck.UseVisualStyleBackColor = true;
            this.alwaysTopCheck.CheckedChanged += new System.EventHandler(this.alwaysTopCheck_CheckedChanged);
            // 
            // CreateFilters
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 424);
            this.Controls.Add(this.alwaysTopCheck);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.deleteRow);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dispProjectName);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.checkTableButton);
            this.Controls.Add(this.clearTable);
            this.Controls.Add(this.loadTable);
            this.Controls.Add(this.saveTable);
            this.Controls.Add(this.dataGridView);
            this.Name = "CreateFilters";
            this.Text = "Create Filter";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button saveTable;
        private System.Windows.Forms.Button loadTable;
        private System.Windows.Forms.Button clearTable;
        private System.Windows.Forms.Button checkTableButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.TextBox dispProjectName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button deleteRow;
        private System.Windows.Forms.DataGridViewTextBoxColumn name;
        private System.Windows.Forms.DataGridViewTextBoxColumn recipient;
        private System.Windows.Forms.DataGridViewTextBoxColumn sender;
        private System.Windows.Forms.DataGridViewTextBoxColumn folderPath;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.CheckBox alwaysTopCheck;
    }
}

