namespace ExcellTimetableOfClasses
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
			this.btnConnect = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.txtNewFile = new System.Windows.Forms.TextBox();
			this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
			this.txtOldFile = new System.Windows.Forms.TextBox();
			this.numEdt = new System.Windows.Forms.NumericUpDown();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.btnGetAll = new System.Windows.Forms.Button();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.richTextBox2 = new System.Windows.Forms.RichTextBox();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.richTextBox3 = new System.Windows.Forms.RichTextBox();
			this.btnDiffer = new System.Windows.Forms.Button();
			this.ds = new System.Data.DataSet();
			this.tblResult = new System.Data.DataTable();
			this.DateKey = new System.Data.DataColumn();
			this.AllOther = new System.Data.DataColumn();
			this.btnWikiStyle = new System.Windows.Forms.Button();
			this.UploadToGCal = new System.Windows.Forms.Button();
			this.chbDate = new System.Windows.Forms.CheckBox();
			this.chbTime = new System.Windows.Forms.CheckBox();
			this.chbSubj = new System.Windows.Forms.CheckBox();
			this.chbTeacher = new System.Windows.Forms.CheckBox();
			this.chbClass = new System.Windows.Forms.CheckBox();
			this.chbOther = new System.Windows.Forms.CheckBox();
			((System.ComponentModel.ISupportInitialize)(this.numEdt)).BeginInit();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.tabPage2.SuspendLayout();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ds)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tblResult)).BeginInit();
			this.SuspendLayout();
			//
			// btnConnect
			//
			this.btnConnect.Location = new System.Drawing.Point(362, 12);
			this.btnConnect.Name = "btnConnect";
			this.btnConnect.Size = new System.Drawing.Size(543, 20);
			this.btnConnect.TabIndex = 0;
			this.btnConnect.Tag = "1";
			this.btnConnect.Text = "Выбрать файлы";
			this.btnConnect.UseVisualStyleBackColor = true;
			this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
			//
			// btnClose
			//
			this.btnClose.Location = new System.Drawing.Point(872, 621);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(150, 49);
			this.btnClose.TabIndex = 1;
			this.btnClose.Tag = "2";
			this.btnClose.Text = "Выход";
			this.btnClose.UseVisualStyleBackColor = true;
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			//
			// richTextBox1
			//
			this.richTextBox1.BackColor = System.Drawing.SystemColors.ScrollBar;
			this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.richTextBox1.Location = new System.Drawing.Point(3, 3);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(996, 444);
			this.richTextBox1.TabIndex = 2;
			this.richTextBox1.Text = "";
			//
			// openFileDialog1
			//
			this.openFileDialog1.Filter = "Excel files|*.xls;*.xlsx";
			this.openFileDialog1.Title = "Выберите новое расписание";
			//
			// txtNewFile
			//
			this.txtNewFile.BackColor = System.Drawing.SystemColors.ScrollBar;
			this.txtNewFile.Location = new System.Drawing.Point(163, 82);
			this.txtNewFile.Name = "txtNewFile";
			this.txtNewFile.Size = new System.Drawing.Size(742, 20);
			this.txtNewFile.TabIndex = 3;
			//
			// openFileDialog2
			//
			this.openFileDialog2.Filter = "Excel files|*.xls;*.xlsx";
			this.openFileDialog2.Title = "А теперь старое";
			//
			// txtOldFile
			//
			this.txtOldFile.BackColor = System.Drawing.SystemColors.ScrollBar;
			this.txtOldFile.Location = new System.Drawing.Point(163, 47);
			this.txtOldFile.Name = "txtOldFile";
			this.txtOldFile.Size = new System.Drawing.Size(742, 20);
			this.txtOldFile.TabIndex = 4;
			//
			// numEdt
			//
			this.numEdt.BackColor = System.Drawing.SystemColors.ScrollBar;
			this.numEdt.Location = new System.Drawing.Point(163, 12);
			this.numEdt.Maximum = new decimal(new int[] {
			15,
			0,
			0,
			0});
			this.numEdt.Name = "numEdt";
			this.numEdt.Size = new System.Drawing.Size(173, 20);
			this.numEdt.TabIndex = 5;
			this.numEdt.Value = new decimal(new int[] {
			2,
			0,
			0,
			0});
			//
			// label2
			//
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(9, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(145, 13);
			this.label2.TabIndex = 7;
			this.label2.Text = "Номер листа для проверки";
			//
			// label3
			//
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(9, 50);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(106, 13);
			this.label3.TabIndex = 8;
			this.label3.Text = "Старое расписание";
			//
			// label4
			//
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(9, 85);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(102, 13);
			this.label4.TabIndex = 9;
			this.label4.Text = "Новое расписание";
			//
			// progressBar1
			//
			this.progressBar1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
			this.progressBar1.Location = new System.Drawing.Point(12, 590);
			this.progressBar1.Maximum = 200;
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(1010, 23);
			this.progressBar1.Step = 2;
			this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
			this.progressBar1.TabIndex = 10;
			//
			// btnGetAll
			//
			this.btnGetAll.Location = new System.Drawing.Point(168, 621);
			this.btnGetAll.Name = "btnGetAll";
			this.btnGetAll.Size = new System.Drawing.Size(150, 49);
			this.btnGetAll.TabIndex = 11;
			this.btnGetAll.Text = "Получить все расписание";
			this.btnGetAll.UseVisualStyleBackColor = true;
			this.btnGetAll.Click += new System.EventHandler(this.btnGetAll_Click);
			//
			// tabControl1
			//
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Controls.Add(this.tabPage3);
			this.tabControl1.Location = new System.Drawing.Point(12, 108);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(1010, 476);
			this.tabControl1.TabIndex = 12;
			//
			// tabPage1
			//
			this.tabPage1.Controls.Add(this.richTextBox1);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage1.Size = new System.Drawing.Size(1002, 450);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Результат сравнения";
			this.tabPage1.UseVisualStyleBackColor = true;
			//
			// tabPage2
			//
			this.tabPage2.Controls.Add(this.richTextBox2);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage2.Size = new System.Drawing.Size(1002, 450);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Полное расписание";
			this.tabPage2.UseVisualStyleBackColor = true;
			//
			// richTextBox2
			//
			this.richTextBox2.BackColor = System.Drawing.SystemColors.ScrollBar;
			this.richTextBox2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.richTextBox2.Location = new System.Drawing.Point(3, 3);
			this.richTextBox2.Name = "richTextBox2";
			this.richTextBox2.Size = new System.Drawing.Size(996, 444);
			this.richTextBox2.TabIndex = 0;
			this.richTextBox2.Text = "";
			//
			// tabPage3
			//
			this.tabPage3.Controls.Add(this.richTextBox3);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage3.Size = new System.Drawing.Size(1002, 450);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Расписание";
			this.tabPage3.UseVisualStyleBackColor = true;
			//
			// richTextBox3
			//
			this.richTextBox3.BackColor = System.Drawing.SystemColors.ScrollBar;
			this.richTextBox3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.richTextBox3.Location = new System.Drawing.Point(3, 3);
			this.richTextBox3.Name = "richTextBox3";
			this.richTextBox3.Size = new System.Drawing.Size(996, 444);
			this.richTextBox3.TabIndex = 0;
			this.richTextBox3.Text = "";
			//
			// btnDiffer
			//
			this.btnDiffer.Location = new System.Drawing.Point(12, 621);
			this.btnDiffer.Name = "btnDiffer";
			this.btnDiffer.Size = new System.Drawing.Size(150, 49);
			this.btnDiffer.TabIndex = 13;
			this.btnDiffer.Text = "Сравнить";
			this.btnDiffer.UseVisualStyleBackColor = true;
			this.btnDiffer.Click += new System.EventHandler(this.btnDiffer_Click);
			//
			// ds
			//
			this.ds.DataSetName = "ds";
			this.ds.Tables.AddRange(new System.Data.DataTable[] {
			this.tblResult});
			//
			// tblResult
			//
			this.tblResult.Columns.AddRange(new System.Data.DataColumn[] {
			this.DateKey,
			this.AllOther});
			this.tblResult.TableName = "ResultByGroup";
			//
			// DateKey
			//
			this.DateKey.ColumnName = "DateKey";
			this.DateKey.DataType = typeof(short);
			this.DateKey.DefaultValue = ((short)(0));
			//
			// AllOther
			//
			this.AllOther.ColumnName = "AllOther";
			//
			// btnWikiStyle
			//
			this.btnWikiStyle.Location = new System.Drawing.Point(324, 621);
			this.btnWikiStyle.Name = "btnWikiStyle";
			this.btnWikiStyle.Size = new System.Drawing.Size(150, 49);
			this.btnWikiStyle.TabIndex = 14;
			this.btnWikiStyle.Text = "в Wiki-стиле";
			this.btnWikiStyle.UseVisualStyleBackColor = true;
			this.btnWikiStyle.Click += new System.EventHandler(this.btnWikiStyle_Click);
			//
			// UploadToGCal
			//
			this.UploadToGCal.Location = new System.Drawing.Point(481, 621);
			this.UploadToGCal.Name = "UploadToGCal";
			this.UploadToGCal.Size = new System.Drawing.Size(173, 49);
			this.UploadToGCal.TabIndex = 19;
			this.UploadToGCal.Text = "Экспортировать в формате календаря";
			this.UploadToGCal.UseVisualStyleBackColor = true;
			//this.UploadToGCal.Click += new System.EventHandler(this.UploadToGCal_Click);
			//
			// chbDate
			//
			this.chbDate.AutoSize = true;
			this.chbDate.Location = new System.Drawing.Point(912, 12);
			this.chbDate.Name = "chbDate";
			this.chbDate.Size = new System.Drawing.Size(52, 17);
			this.chbDate.TabIndex = 20;
			this.chbDate.Text = "Дата";
			this.chbDate.UseVisualStyleBackColor = true;
			//
			// chbTime
			//
			this.chbTime.AutoSize = true;
			this.chbTime.Location = new System.Drawing.Point(912, 26);
			this.chbTime.Name = "chbTime";
			this.chbTime.Size = new System.Drawing.Size(59, 17);
			this.chbTime.TabIndex = 21;
			this.chbTime.Text = "Время";
			this.chbTime.UseVisualStyleBackColor = true;
			//
			// chbSubj
			//
			this.chbSubj.AutoSize = true;
			this.chbSubj.Location = new System.Drawing.Point(912, 40);
			this.chbSubj.Name = "chbSubj";
			this.chbSubj.Size = new System.Drawing.Size(71, 17);
			this.chbSubj.TabIndex = 22;
			this.chbSubj.Text = "Предмет";
			this.chbSubj.UseVisualStyleBackColor = true;
			//
			// chbTeacher
			//
			this.chbTeacher.AutoSize = true;
			this.chbTeacher.Location = new System.Drawing.Point(912, 54);
			this.chbTeacher.Name = "chbTeacher";
			this.chbTeacher.Size = new System.Drawing.Size(105, 17);
			this.chbTeacher.TabIndex = 23;
			this.chbTeacher.Text = "Преподаватель";
			this.chbTeacher.UseVisualStyleBackColor = true;
			//
			// chbClass
			//
			this.chbClass.AutoSize = true;
			this.chbClass.Location = new System.Drawing.Point(912, 68);
			this.chbClass.Name = "chbClass";
			this.chbClass.Size = new System.Drawing.Size(79, 17);
			this.chbClass.TabIndex = 24;
			this.chbClass.Text = "Аудитория";
			this.chbClass.UseVisualStyleBackColor = true;
			//
			// chbOther
			//
			this.chbOther.AutoSize = true;
			this.chbOther.Location = new System.Drawing.Point(912, 82);
			this.chbOther.Name = "chbOther";
			this.chbOther.Size = new System.Drawing.Size(89, 17);
			this.chbOther.TabIndex = 25;
			this.chbOther.Text = "Примечание";
			this.chbOther.UseVisualStyleBackColor = true;
			//
			// Form1
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.ButtonShadow;
			this.ClientSize = new System.Drawing.Size(1034, 682);
			this.Controls.Add(this.chbOther);
			this.Controls.Add(this.chbClass);
			this.Controls.Add(this.chbTeacher);
			this.Controls.Add(this.chbSubj);
			this.Controls.Add(this.chbTime);
			this.Controls.Add(this.chbDate);
			this.Controls.Add(this.UploadToGCal);
			this.Controls.Add(this.btnWikiStyle);
			this.Controls.Add(this.btnDiffer);
			this.Controls.Add(this.btnGetAll);
			this.Controls.Add(this.progressBar1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.numEdt);
			this.Controls.Add(this.txtOldFile);
			this.Controls.Add(this.txtNewFile);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.btnConnect);
			this.Controls.Add(this.tabControl1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Сравнение расписания";
			((System.ComponentModel.ISupportInitialize)(this.numEdt)).EndInit();
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ds)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tblResult)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btnConnect;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.RichTextBox richTextBox1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.TextBox txtNewFile;
		private System.Windows.Forms.OpenFileDialog openFileDialog2;
		private System.Windows.Forms.TextBox txtOldFile;
		private System.Windows.Forms.NumericUpDown numEdt;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.Button btnGetAll;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.RichTextBox richTextBox2;
		private System.Windows.Forms.Button btnDiffer;
		private System.Data.DataSet ds;
		private System.Data.DataTable tblResult;
		private System.Data.DataColumn DateKey;
		private System.Data.DataColumn AllOther;
		private System.Windows.Forms.Button btnWikiStyle;
		private System.Windows.Forms.Button UploadToGCal;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.RichTextBox richTextBox3;
		private System.Windows.Forms.CheckBox chbDate;
		private System.Windows.Forms.CheckBox chbTime;
		private System.Windows.Forms.CheckBox chbSubj;
		private System.Windows.Forms.CheckBox chbTeacher;
		private System.Windows.Forms.CheckBox chbClass;
		private System.Windows.Forms.CheckBox chbOther;
	}
}

