namespace DMS.Utilities.ExcelExport
{
    partial class Exporter
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbxConnection = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.numPageSize = new System.Windows.Forms.NumericUpDown();
            this.ckbBreak = new System.Windows.Forms.CheckBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.ckbStores = new System.Windows.Forms.CheckBox();
            this.tbxStores = new System.Windows.Forms.RichTextBox();
            this.tbxTimeout = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numPageSize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbxTimeout)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "SQL Connection";
            // 
            // tbxConnection
            // 
            this.tbxConnection.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbxConnection.Location = new System.Drawing.Point(108, 12);
            this.tbxConnection.Name = "tbxConnection";
            this.tbxConnection.Size = new System.Drawing.Size(313, 20);
            this.tbxConnection.TabIndex = 1;
            this.tbxConnection.Text = "data source=10.88.17.51\\SQLSERVER_PEPSI;initial catalog=EMSDW;persist security in" +
    "fo=True;user id=u_dms;password=1qaZ2wsX;";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Command";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Row/Sheet";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.numPageSize);
            this.groupBox1.Controls.Add(this.ckbBreak);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(12, 185);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(409, 54);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Export Configurations";
            // 
            // numPageSize
            // 
            this.numPageSize.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.numPageSize.Location = new System.Drawing.Point(96, 20);
            this.numPageSize.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.numPageSize.Name = "numPageSize";
            this.numPageSize.Size = new System.Drawing.Size(222, 20);
            this.numPageSize.TabIndex = 6;
            // 
            // ckbBreak
            // 
            this.ckbBreak.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.ckbBreak.AutoSize = true;
            this.ckbBreak.Location = new System.Drawing.Point(324, 21);
            this.ckbBreak.Name = "ckbBreak";
            this.ckbBreak.Size = new System.Drawing.Size(79, 17);
            this.ckbBreak.TabIndex = 4;
            this.ckbBreak.Text = "Multi-Sheet";
            this.ckbBreak.UseVisualStyleBackColor = true;
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.Location = new System.Drawing.Point(346, 245);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 5;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // ckbStores
            // 
            this.ckbStores.AutoSize = true;
            this.ckbStores.Location = new System.Drawing.Point(12, 83);
            this.ckbStores.Name = "ckbStores";
            this.ckbStores.Size = new System.Drawing.Size(56, 17);
            this.ckbStores.TabIndex = 4;
            this.ckbStores.Text = "Stores";
            this.ckbStores.UseVisualStyleBackColor = true;
            // 
            // tbxStores
            // 
            this.tbxStores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbxStores.Location = new System.Drawing.Point(108, 64);
            this.tbxStores.Name = "tbxStores";
            this.tbxStores.Size = new System.Drawing.Size(313, 115);
            this.tbxStores.TabIndex = 6;
            this.tbxStores.Text = "Select top 10 * from dimproduct";
            // 
            // tbxTimeout
            // 
            this.tbxTimeout.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbxTimeout.Location = new System.Drawing.Point(108, 38);
            this.tbxTimeout.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.tbxTimeout.Name = "tbxTimeout";
            this.tbxTimeout.Size = new System.Drawing.Size(313, 20);
            this.tbxTimeout.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 40);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Timeout";
            // 
            // Exporter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(433, 280);
            this.Controls.Add(this.tbxTimeout);
            this.Controls.Add(this.tbxStores);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.ckbStores);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.tbxConnection);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MinimumSize = new System.Drawing.Size(300, 208);
            this.Name = "Exporter";
            this.Text = "Exporter";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numPageSize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbxTimeout)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbxConnection;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox ckbBreak;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.NumericUpDown numPageSize;
        private System.Windows.Forms.CheckBox ckbStores;
        private System.Windows.Forms.RichTextBox tbxStores;
        private System.Windows.Forms.NumericUpDown tbxTimeout;
        private System.Windows.Forms.Label label4;
    }
}

