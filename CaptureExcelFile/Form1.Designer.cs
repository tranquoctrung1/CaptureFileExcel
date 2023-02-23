namespace CaptureExcelFile
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button2 = new System.Windows.Forms.Button();
            this.btn_browserFile = new System.Windows.Forms.Button();
            this.lbTitle = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtProductId = new System.Windows.Forms.TextBox();
            this.btnChoosePathToSaveImage = new System.Windows.Forms.Button();
            this.txtDescription = new System.Windows.Forms.RichTextBox();
            this.txtDate = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.ckSplit = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(410, 160);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "5/ Thực thi";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btn_browserFile
            // 
            this.btn_browserFile.Location = new System.Drawing.Point(26, 111);
            this.btn_browserFile.Name = "btn_browserFile";
            this.btn_browserFile.Size = new System.Drawing.Size(75, 23);
            this.btn_browserFile.TabIndex = 3;
            this.btn_browserFile.Text = "1/ Chọn file";
            this.btn_browserFile.UseVisualStyleBackColor = true;
            this.btn_browserFile.Click += new System.EventHandler(this.btn_browserFile_Click);
            // 
            // lbTitle
            // 
            this.lbTitle.AutoSize = true;
            this.lbTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTitle.Location = new System.Drawing.Point(365, 27);
            this.lbTitle.Name = "lbTitle";
            this.lbTitle.Size = new System.Drawing.Size(301, 25);
            this.lbTitle.TabIndex = 4;
            this.lbTitle.Text = "Lọc file excel và chụp hình ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(280, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "3/ Mã hàng hóa";
            // 
            // txtProductId
            // 
            this.txtProductId.Location = new System.Drawing.Point(370, 108);
            this.txtProductId.Name = "txtProductId";
            this.txtProductId.Size = new System.Drawing.Size(212, 20);
            this.txtProductId.TabIndex = 6;
            // 
            // btnChoosePathToSaveImage
            // 
            this.btnChoosePathToSaveImage.Location = new System.Drawing.Point(126, 109);
            this.btnChoosePathToSaveImage.Name = "btnChoosePathToSaveImage";
            this.btnChoosePathToSaveImage.Size = new System.Drawing.Size(136, 23);
            this.btnChoosePathToSaveImage.TabIndex = 7;
            this.btnChoosePathToSaveImage.Text = "2/ Chọn nơi lưu ảnh";
            this.btnChoosePathToSaveImage.UseVisualStyleBackColor = true;
            this.btnChoosePathToSaveImage.Click += new System.EventHandler(this.btnChoosePathToSaveImage_Click);
            // 
            // txtDescription
            // 
            this.txtDescription.Location = new System.Drawing.Point(26, 210);
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(908, 266);
            this.txtDescription.TabIndex = 10;
            this.txtDescription.Text = "";
            // 
            // txtDate
            // 
            this.txtDate.Location = new System.Drawing.Point(734, 109);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(200, 20);
            this.txtDate.TabIndex = 11;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(656, 111);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = "4/ Chọn ngày";
            // 
            // ckSplit
            // 
            this.ckSplit.AutoSize = true;
            this.ckSplit.Location = new System.Drawing.Point(588, 111);
            this.ckSplit.Name = "ckSplit";
            this.ckSplit.Size = new System.Drawing.Size(51, 17);
            this.ckSplit.TabIndex = 14;
            this.ckSplit.Text = "Tách";
            this.ckSplit.UseVisualStyleBackColor = true;
            this.ckSplit.CheckedChanged += new System.EventHandler(this.ckSplit_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(965, 506);
            this.Controls.Add(this.ckSplit);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.txtDescription);
            this.Controls.Add(this.btnChoosePathToSaveImage);
            this.Controls.Add(this.txtProductId);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lbTitle);
            this.Controls.Add(this.btn_browserFile);
            this.Controls.Add(this.button2);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NHẤT NAM FOOD";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btn_browserFile;
        private System.Windows.Forms.Label lbTitle;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtProductId;
        private System.Windows.Forms.Button btnChoosePathToSaveImage;
        private System.Windows.Forms.RichTextBox txtDescription;
        private System.Windows.Forms.DateTimePicker txtDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox ckSplit;
    }
}

