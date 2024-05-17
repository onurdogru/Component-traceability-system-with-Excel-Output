
namespace CANIAS_Izlenebilirlik
{
    partial class CardList
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
            this.txtCode = new System.Windows.Forms.TextBox();
            this.dataGridViewCL = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMaterialName = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.btnOpenSerialPortAlarm = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnWarningLamb = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtPublicFolder = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCL)).BeginInit();
            this.SuspendLayout();
            // 
            // txtCode
            // 
            this.txtCode.BackColor = System.Drawing.Color.White;
            this.txtCode.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtCode.Location = new System.Drawing.Point(907, 53);
            this.txtCode.Name = "txtCode";
            this.txtCode.Size = new System.Drawing.Size(204, 32);
            this.txtCode.TabIndex = 1;
            this.txtCode.TextChanged += new System.EventHandler(this.txtCode_TextChanged);
            // 
            // dataGridViewCL
            // 
            this.dataGridViewCL.AllowUserToAddRows = false;
            this.dataGridViewCL.AllowUserToDeleteRows = false;
            this.dataGridViewCL.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewCL.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridViewCL.ColumnHeadersHeight = 29;
            this.dataGridViewCL.EnableHeadersVisualStyles = false;
            this.dataGridViewCL.Location = new System.Drawing.Point(1, 1);
            this.dataGridViewCL.Name = "dataGridViewCL";
            this.dataGridViewCL.ReadOnly = true;
            this.dataGridViewCL.RowHeadersVisible = false;
            this.dataGridViewCL.RowHeadersWidth = 51;
            this.dataGridViewCL.RowTemplate.Height = 24;
            this.dataGridViewCL.Size = new System.Drawing.Size(857, 591);
            this.dataGridViewCL.TabIndex = 3;
            this.dataGridViewCL.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewCL_CellClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(927, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 23);
            this.label1.TabIndex = 4;
            this.label1.Text = "Malzeme Kodu";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(927, 124);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(145, 23);
            this.label2.TabIndex = 6;
            this.label2.Text = "Malzeme İsmi";
            // 
            // txtMaterialName
            // 
            this.txtMaterialName.BackColor = System.Drawing.Color.White;
            this.txtMaterialName.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtMaterialName.Location = new System.Drawing.Point(907, 159);
            this.txtMaterialName.Name = "txtMaterialName";
            this.txtMaterialName.Size = new System.Drawing.Size(204, 32);
            this.txtMaterialName.TabIndex = 5;
            this.txtMaterialName.TextChanged += new System.EventHandler(this.txtMaterialName_TextChanged);
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(926, 432);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(200, 31);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.Visible = false;
            // 
            // btnOpenSerialPortAlarm
            // 
            this.btnOpenSerialPortAlarm.BackColor = System.Drawing.Color.Olive;
            this.btnOpenSerialPortAlarm.FlatAppearance.BorderSize = 3;
            this.btnOpenSerialPortAlarm.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenSerialPortAlarm.Font = new System.Drawing.Font("Century Gothic", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnOpenSerialPortAlarm.ForeColor = System.Drawing.Color.White;
            this.btnOpenSerialPortAlarm.Location = new System.Drawing.Point(926, 485);
            this.btnOpenSerialPortAlarm.Name = "btnOpenSerialPortAlarm";
            this.btnOpenSerialPortAlarm.Size = new System.Drawing.Size(95, 84);
            this.btnOpenSerialPortAlarm.TabIndex = 8;
            this.btnOpenSerialPortAlarm.Text = "ON";
            this.btnOpenSerialPortAlarm.UseVisualStyleBackColor = false;
            this.btnOpenSerialPortAlarm.Visible = false;
            this.btnOpenSerialPortAlarm.Click += new System.EventHandler(this.btnOpenSerialPortAlarm_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(946, 396);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(137, 23);
            this.label3.TabIndex = 9;
            this.label3.Text = "İkaz Lambası";
            this.label3.Visible = false;
            // 
            // btnWarningLamb
            // 
            this.btnWarningLamb.BackColor = System.Drawing.Color.Transparent;
            this.btnWarningLamb.BackgroundImage = global::CANIAS_Izlenebilirlik.Properties.Resources.WarningLamb;
            this.btnWarningLamb.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnWarningLamb.FlatAppearance.BorderSize = 0;
            this.btnWarningLamb.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnWarningLamb.Location = new System.Drawing.Point(994, 313);
            this.btnWarningLamb.Name = "btnWarningLamb";
            this.btnWarningLamb.Size = new System.Drawing.Size(51, 76);
            this.btnWarningLamb.TabIndex = 10;
            this.btnWarningLamb.UseVisualStyleBackColor = false;
            this.btnWarningLamb.Click += new System.EventHandler(this.btnWarningLamb_Click);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.Green;
            this.btnSave.FlatAppearance.BorderSize = 3;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Font = new System.Drawing.Font("Century Gothic", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnSave.ForeColor = System.Drawing.Color.White;
            this.btnSave.Location = new System.Drawing.Point(1031, 485);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(95, 84);
            this.btnSave.TabIndex = 11;
            this.btnSave.Text = "SAVE";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Visible = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // txtPublicFolder
            // 
            this.txtPublicFolder.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.txtPublicFolder.Location = new System.Drawing.Point(907, 255);
            this.txtPublicFolder.Name = "txtPublicFolder";
            this.txtPublicFolder.Size = new System.Drawing.Size(204, 26);
            this.txtPublicFolder.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(903, 226);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(128, 23);
            this.label4.TabIndex = 13;
            this.label4.Text = "Ortak Klasör";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.checkBox1.ForeColor = System.Drawing.Color.Green;
            this.checkBox1.Location = new System.Drawing.Point(1045, 226);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(81, 23);
            this.checkBox1.TabIndex = 14;
            this.checkBox1.Text = "Kullan";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // CardList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.ClientSize = new System.Drawing.Size(1156, 591);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtPublicFolder);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnWarningLamb);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnOpenSerialPortAlarm);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtMaterialName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridViewCL);
            this.Controls.Add(this.txtCode);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "CardList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "\"\"";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CardList_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CardList_FormClosed);
            this.Load += new System.EventHandler(this.CardList_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCL)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtCode;
        private System.Windows.Forms.DataGridView dataGridViewCL;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMaterialName;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button btnOpenSerialPortAlarm;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnWarningLamb;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtPublicFolder;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}