namespace DocumentsModule
{
    partial class ForMeToSign
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridViewDocs = new System.Windows.Forms.DataGridView();
            this.docIdColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.creatorColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dateColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.signColumn = new System.Windows.Forms.DataGridViewButtonColumn();
            this.publishColumn = new System.Windows.Forms.DataGridViewButtonColumn();
            this.returnColumn = new System.Windows.Forms.DataGridViewButtonColumn();
            this.documentsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.documentsSignTableAdapter = new DocumentsModule.MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter();
            this.mantakDBDataSetDocuments = new DocumentsModule.MantakDBDataSetDocuments();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDocs)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.documentsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mantakDBDataSetDocuments)).BeginInit();
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
            this.creatorColumn,
            this.dateColumn,
            this.signColumn,
            this.publishColumn,
            this.returnColumn});
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
            this.dataGridViewDocs.Size = new System.Drawing.Size(974, 460);
            this.dataGridViewDocs.TabIndex = 53;
            this.dataGridViewDocs.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewDocs_CellContentClick);
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
            this.nameColumn.Width = 375;
            // 
            // creatorColumn
            // 
            this.creatorColumn.HeaderText = "יוצר מסמך";
            this.creatorColumn.Name = "creatorColumn";
            this.creatorColumn.ReadOnly = true;
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
            // signColumn
            // 
            this.signColumn.HeaderText = "חתימה";
            this.signColumn.Name = "signColumn";
            this.signColumn.ReadOnly = true;
            this.signColumn.Width = 90;
            // 
            // publishColumn
            // 
            this.publishColumn.HeaderText = "הפצה";
            this.publishColumn.Name = "publishColumn";
            this.publishColumn.ReadOnly = true;
            this.publishColumn.Width = 65;
            // 
            // returnColumn
            // 
            this.returnColumn.HeaderText = "החזר להמשך עבודה";
            this.returnColumn.Name = "returnColumn";
            this.returnColumn.ReadOnly = true;
            this.returnColumn.Width = 130;
            // 
            // documentsBindingSource
            // 
            this.documentsBindingSource.DataMember = "documentsSign";
            // 
            // documentsSignTableAdapter
            // 
            this.documentsSignTableAdapter.ClearBeforeFill = true;
            // 
            // mantakDBDataSetDocuments
            // 
            this.mantakDBDataSetDocuments.DataSetName = "MantakDBDataSetDocuments";
            this.mantakDBDataSetDocuments.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // ForMeToSign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(996, 484);
            this.Controls.Add(this.dataGridViewDocs);
            this.MaximizeBox = false;
            this.Name = "ForMeToSign";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "מסמכים לחתימתי";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ForMeToSign_FormClosed);
            this.Load += new System.EventHandler(this.ForMeToSign_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDocs)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.documentsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mantakDBDataSetDocuments)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewDocs;
        private MantakDBDataSetDocuments mantakDBDataSetDocuments;
        private System.Windows.Forms.BindingSource documentsBindingSource;
        private MantakDBDataSetDocumentsTableAdapters.documentsSignTableAdapter documentsSignTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn docIdColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nameColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn creatorColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dateColumn;
        private System.Windows.Forms.DataGridViewButtonColumn signColumn;
        private System.Windows.Forms.DataGridViewButtonColumn publishColumn;
        private System.Windows.Forms.DataGridViewButtonColumn returnColumn;
    }
}