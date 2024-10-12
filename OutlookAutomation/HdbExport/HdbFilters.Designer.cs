namespace OutlookAutomation
{
    partial class HdbFilters
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
            this.recipientGridView = new System.Windows.Forms.DataGridView();
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.label2 = new System.Windows.Forms.Label();
            this.senderGridView = new System.Windows.Forms.DataGridView();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.subjectGridView = new System.Windows.Forms.DataGridView();
            this.subjectText = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.subjectNumCharacters = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.externalReferenceName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.internalSender = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.recipientGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.senderGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.subjectGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // recipientGridView
            // 
            this.recipientGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.recipientGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.recipientGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.recipientGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.email,
            this.externalReferenceName});
            this.recipientGridView.Location = new System.Drawing.Point(6, 42);
            this.recipientGridView.Margin = new System.Windows.Forms.Padding(6);
            this.recipientGridView.Name = "recipientGridView";
            this.recipientGridView.RowHeadersWidth = 72;
            this.recipientGridView.Size = new System.Drawing.Size(1440, 273);
            this.recipientGridView.TabIndex = 2;
            // 
            // saveTable
            // 
            this.saveTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.saveTable.Location = new System.Drawing.Point(546, 1073);
            this.saveTable.Margin = new System.Windows.Forms.Padding(6);
            this.saveTable.Name = "saveTable";
            this.saveTable.Size = new System.Drawing.Size(183, 48);
            this.saveTable.TabIndex = 6;
            this.saveTable.Text = "Export Table";
            this.saveTable.UseVisualStyleBackColor = true;
            this.saveTable.Click += new System.EventHandler(this.saveTable_Click);
            // 
            // loadTable
            // 
            this.loadTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.loadTable.Location = new System.Drawing.Point(739, 1073);
            this.loadTable.Margin = new System.Windows.Forms.Padding(6);
            this.loadTable.Name = "loadTable";
            this.loadTable.Size = new System.Drawing.Size(165, 48);
            this.loadTable.TabIndex = 7;
            this.loadTable.Text = "Load Table";
            this.loadTable.UseVisualStyleBackColor = true;
            this.loadTable.Click += new System.EventHandler(this.loadTable_Click);
            // 
            // clearTable
            // 
            this.clearTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.clearTable.Location = new System.Drawing.Point(369, 1073);
            this.clearTable.Margin = new System.Windows.Forms.Padding(6);
            this.clearTable.Name = "clearTable";
            this.clearTable.Size = new System.Drawing.Size(165, 48);
            this.clearTable.TabIndex = 5;
            this.clearTable.Text = "ClearTable";
            this.clearTable.UseVisualStyleBackColor = true;
            this.clearTable.Click += new System.EventHandler(this.clearTable_Click);
            // 
            // checkTableButton
            // 
            this.checkTableButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkTableButton.Location = new System.Drawing.Point(15, 1073);
            this.checkTableButton.Margin = new System.Windows.Forms.Padding(6);
            this.checkTableButton.Name = "checkTableButton";
            this.checkTableButton.Size = new System.Drawing.Size(165, 48);
            this.checkTableButton.TabIndex = 3;
            this.checkTableButton.Text = "Check Table";
            this.checkTableButton.UseVisualStyleBackColor = true;
            this.checkTableButton.Click += new System.EventHandler(this.checkTableButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Location = new System.Drawing.Point(1296, 1073);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(6);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(165, 48);
            this.cancelButton.TabIndex = 9;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // dispProjectName
            // 
            this.dispProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dispProjectName.Location = new System.Drawing.Point(28, 54);
            this.dispProjectName.Margin = new System.Windows.Forms.Padding(6);
            this.dispProjectName.Name = "dispProjectName";
            this.dispProjectName.Size = new System.Drawing.Size(1423, 29);
            this.dispProjectName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(22, 17);
            this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(456, 31);
            this.label1.TabIndex = 7;
            this.label1.Text = "Project Name:";
            // 
            // deleteRow
            // 
            this.deleteRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.deleteRow.Location = new System.Drawing.Point(192, 1073);
            this.deleteRow.Margin = new System.Windows.Forms.Padding(6);
            this.deleteRow.Name = "deleteRow";
            this.deleteRow.Size = new System.Drawing.Size(165, 48);
            this.deleteRow.TabIndex = 4;
            this.deleteRow.Text = "Delete Row";
            this.deleteRow.UseVisualStyleBackColor = true;
            this.deleteRow.Click += new System.EventHandler(this.deleteRow_Click);
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.Location = new System.Drawing.Point(1123, 1073);
            this.okButton.Margin = new System.Windows.Forms.Padding(6);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(165, 48);
            this.okButton.TabIndex = 8;
            this.okButton.Text = "Ok";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // alwaysTopCheck
            // 
            this.alwaysTopCheck.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.alwaysTopCheck.AutoSize = true;
            this.alwaysTopCheck.Location = new System.Drawing.Point(1290, 7);
            this.alwaysTopCheck.Margin = new System.Windows.Forms.Padding(4);
            this.alwaysTopCheck.Name = "alwaysTopCheck";
            this.alwaysTopCheck.Size = new System.Drawing.Size(173, 29);
            this.alwaysTopCheck.TabIndex = 10;
            this.alwaysTopCheck.Text = "Always On Top";
            this.alwaysTopCheck.UseVisualStyleBackColor = true;
            this.alwaysTopCheck.CheckedChanged += new System.EventHandler(this.alwaysTopCheck_CheckedChanged);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(12, 92);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.label2);
            this.splitContainer1.Panel1.Controls.Add(this.senderGridView);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(1452, 968);
            this.splitContainer1.SplitterDistance = 317;
            this.splitContainer1.TabIndex = 12;
            this.splitContainer1.TabStop = false;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(6, 5);
            this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(456, 31);
            this.label2.TabIndex = 13;
            this.label2.Text = "Sender:";
            // 
            // senderGridView
            // 
            this.senderGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.senderGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.senderGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.senderGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.internalSender});
            this.senderGridView.Location = new System.Drawing.Point(6, 42);
            this.senderGridView.Margin = new System.Windows.Forms.Padding(6);
            this.senderGridView.Name = "senderGridView";
            this.senderGridView.RowHeadersWidth = 72;
            this.senderGridView.Size = new System.Drawing.Size(1440, 267);
            this.senderGridView.TabIndex = 12;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.label3);
            this.splitContainer2.Panel1.Controls.Add(this.recipientGridView);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.label4);
            this.splitContainer2.Panel2.Controls.Add(this.subjectGridView);
            this.splitContainer2.Size = new System.Drawing.Size(1452, 647);
            this.splitContainer2.SplitterDistance = 321;
            this.splitContainer2.TabIndex = 0;
            this.splitContainer2.TabStop = false;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(6, 5);
            this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(456, 31);
            this.label3.TabIndex = 14;
            this.label3.Text = "External:";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(6, 5);
            this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(456, 31);
            this.label4.TabIndex = 15;
            this.label4.Text = "Subject:";
            // 
            // subjectGridView
            // 
            this.subjectGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.subjectGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.subjectGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.subjectGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.subjectText,
            this.subjectNumCharacters});
            this.subjectGridView.Location = new System.Drawing.Point(6, 42);
            this.subjectGridView.Margin = new System.Windows.Forms.Padding(6);
            this.subjectGridView.Name = "subjectGridView";
            this.subjectGridView.RowHeadersWidth = 72;
            this.subjectGridView.Size = new System.Drawing.Size(1440, 274);
            this.subjectGridView.TabIndex = 3;
            // 
            // subjectText
            // 
            this.subjectText.HeaderText = "Subject Text";
            this.subjectText.MinimumWidth = 9;
            this.subjectText.Name = "subjectText";
            this.subjectText.ToolTipText = "Start of string to search in the subject. If subject contains string, it will rep" +
    "lace the subject folder created with this string + number of characters defined." +
    " ";
            // 
            // subjectNumCharacters
            // 
            this.subjectNumCharacters.HeaderText = "Number of Characters";
            this.subjectNumCharacters.MinimumWidth = 9;
            this.subjectNumCharacters.Name = "subjectNumCharacters";
            this.subjectNumCharacters.ToolTipText = "Number of characters to be considered as new subject folder name. If set to 0, it" +
    " will use subject text.";
            // 
            // email
            // 
            this.email.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.email.HeaderText = "External Party (Email/Domain)";
            this.email.MinimumWidth = 9;
            this.email.Name = "email";
            // 
            // externalReferenceName
            // 
            this.externalReferenceName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.externalReferenceName.HeaderText = "Reference Name";
            this.externalReferenceName.MinimumWidth = 9;
            this.externalReferenceName.Name = "externalReferenceName";
            this.externalReferenceName.ToolTipText = "e.g. Archi, QP(S), Contractor";
            // 
            // internalSender
            // 
            this.internalSender.HeaderText = "Internal Sender (Name)";
            this.internalSender.MinimumWidth = 9;
            this.internalSender.Name = "internalSender";
            this.internalSender.ToolTipText = "Full name including brackets. Not case sensitive.";
            // 
            // HdbFilters
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1476, 1136);
            this.Controls.Add(this.splitContainer1);
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
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MinimumSize = new System.Drawing.Size(1300, 64);
            this.Name = "HdbFilters";
            this.Text = "Create Filter";
            ((System.ComponentModel.ISupportInitialize)(this.recipientGridView)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.senderGridView)).EndInit();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.subjectGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView recipientGridView;
        private System.Windows.Forms.Button saveTable;
        private System.Windows.Forms.Button loadTable;
        private System.Windows.Forms.Button clearTable;
        private System.Windows.Forms.Button checkTableButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.TextBox dispProjectName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button deleteRow;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.CheckBox alwaysTopCheck;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView senderGridView;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView subjectGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn subjectText;
        private System.Windows.Forms.DataGridViewTextBoxColumn subjectNumCharacters;
        private System.Windows.Forms.DataGridViewTextBoxColumn email;
        private System.Windows.Forms.DataGridViewTextBoxColumn externalReferenceName;
        private System.Windows.Forms.DataGridViewTextBoxColumn internalSender;
    }
}

