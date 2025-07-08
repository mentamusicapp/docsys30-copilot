namespace DocumentsModule
{
    partial class MyDocsStatus
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridViewDocs = new System.Windows.Forms.DataGridView();
            this.docIdColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.senderColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dateColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.statusColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDocs)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewDocs
            // 
            this.dataGridViewDocs.AllowUserToAddRows = false;
            this.dataGridViewDocs.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewDocs.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewDocs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDocs.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.docIdColumn,
            this.nameColumn,
            this.senderColumn,
            this.dateColumn,
            this.statusColumn});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewDocs.DefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewDocs.Location = new System.Drawing.Point(12, 12);
            this.dataGridViewDocs.MultiSelect = false;
            this.dataGridViewDocs.Name = "dataGridViewDocs";
            this.dataGridViewDocs.ReadOnly = true;
            this.dataGridViewDocs.RowHeadersVisible = false;
            this.dataGridViewDocs.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridViewDocs.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewDocs.Size = new System.Drawing.Size(755, 460);
            this.dataGridViewDocs.TabIndex = 55;
            this.dataGridViewDocs.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewDocs_CellDoubleClick);
            // 
            // docIdColumn
            // 
            this.docIdColumn.HeaderText = "שוטף";
            this.docIdColumn.Name = "docIdColumn";
            this.docIdColumn.ReadOnly = true;
            this.docIdColumn.Width = 90;
            // 
            // nameColumn
            // 
            this.nameColumn.HeaderText = "נושא";
            this.nameColumn.Name = "nameColumn";
            this.nameColumn.ReadOnly = true;
            this.nameColumn.Width = 305;
            // 
            // senderColumn
            // 
            this.senderColumn.HeaderText = "משתמש חותם";
            this.senderColumn.Name = "senderColumn";
            this.senderColumn.ReadOnly = true;
            this.senderColumn.Width = 105;
            // 
            // dateColumn
            // 
            dataGridViewCellStyle2.Format = "d";
            dataGridViewCellStyle2.NullValue = null;
            this.dateColumn.DefaultCellStyle = dataGridViewCellStyle2;
            this.dateColumn.HeaderText = "תאריך יצירה";
            this.dateColumn.Name = "dateColumn";
            this.dateColumn.ReadOnly = true;
            // 
            // statusColumn
            // 
            this.statusColumn.HeaderText = "סטטוס";
            this.statusColumn.Name = "statusColumn";
            this.statusColumn.ReadOnly = true;
            this.statusColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.statusColumn.Width = 133;
            // 
            // MyDocsStatus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(781, 483);
            this.Controls.Add(this.dataGridViewDocs);
            this.Name = "MyDocsStatus";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "סטטוס המסמכים שלי";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MyDocsStatus_FormClosed);
            this.Load += new System.EventHandler(this.MyDocsStatus_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDocs)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewDocs;
        private System.Windows.Forms.DataGridViewTextBoxColumn docIdColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nameColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn senderColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dateColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn statusColumn;
    }
}