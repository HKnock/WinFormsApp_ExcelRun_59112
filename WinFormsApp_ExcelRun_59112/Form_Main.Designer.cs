
namespace WinFormsApp_ExcelRun_59112
{
    partial class form_Main
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.listBox_InputCell = new System.Windows.Forms.ListBox();
            this.listBox_InputData = new System.Windows.Forms.ListBox();
            this.button_WriteColumn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_InputFormula = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox_CellFormula = new System.Windows.Forms.TextBox();
            this.textBox_CellData = new System.Windows.Forms.TextBox();
            this.button_WriteFormula = new System.Windows.Forms.Button();
            this.button_ReadCell = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox_ResultFormula = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_ResultData = new System.Windows.Forms.TextBox();
            this.button_RunExcel = new System.Windows.Forms.Button();
            this.button_HideExcel = new System.Windows.Forms.Button();
            this.button_ShowExcel = new System.Windows.Forms.Button();
            this.button_Close_Excel = new System.Windows.Forms.Button();
            this.button_Exit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 30);
            this.label1.TabIndex = 0;
            this.label1.Text = "Номер ячеек и \r\nданные для записи";
            // 
            // listBox_InputCell
            // 
            this.listBox_InputCell.FormattingEnabled = true;
            this.listBox_InputCell.ItemHeight = 15;
            this.listBox_InputCell.Location = new System.Drawing.Point(12, 46);
            this.listBox_InputCell.Name = "listBox_InputCell";
            this.listBox_InputCell.Size = new System.Drawing.Size(50, 169);
            this.listBox_InputCell.TabIndex = 1;
            // 
            // listBox_InputData
            // 
            this.listBox_InputData.FormattingEnabled = true;
            this.listBox_InputData.ItemHeight = 15;
            this.listBox_InputData.Location = new System.Drawing.Point(69, 46);
            this.listBox_InputData.Name = "listBox_InputData";
            this.listBox_InputData.Size = new System.Drawing.Size(50, 169);
            this.listBox_InputData.TabIndex = 2;
            // 
            // button_WriteColumn
            // 
            this.button_WriteColumn.Location = new System.Drawing.Point(12, 221);
            this.button_WriteColumn.Name = "button_WriteColumn";
            this.button_WriteColumn.Size = new System.Drawing.Size(221, 39);
            this.button_WriteColumn.TabIndex = 3;
            this.button_WriteColumn.Text = "Записать данные в ячейки в столбец";
            this.button_WriteColumn.UseVisualStyleBackColor = true;
            this.button_WriteColumn.Click += new System.EventHandler(this.button_WriteColumn_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(154, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "Формула";
            // 
            // textBox_InputFormula
            // 
            this.textBox_InputFormula.Location = new System.Drawing.Point(218, 24);
            this.textBox_InputFormula.Name = "textBox_InputFormula";
            this.textBox_InputFormula.Size = new System.Drawing.Size(414, 23);
            this.textBox_InputFormula.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(136, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 30);
            this.label3.TabIndex = 6;
            this.label3.Text = "Ячейка для \r\nформулы";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(136, 140);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 30);
            this.label4.TabIndex = 7;
            this.label4.Text = "Ячейка \r\nданных";
            // 
            // textBox_CellFormula
            // 
            this.textBox_CellFormula.Location = new System.Drawing.Point(218, 71);
            this.textBox_CellFormula.Name = "textBox_CellFormula";
            this.textBox_CellFormula.Size = new System.Drawing.Size(46, 23);
            this.textBox_CellFormula.TabIndex = 8;
            // 
            // textBox_CellData
            // 
            this.textBox_CellData.Location = new System.Drawing.Point(218, 147);
            this.textBox_CellData.Name = "textBox_CellData";
            this.textBox_CellData.Size = new System.Drawing.Size(46, 23);
            this.textBox_CellData.TabIndex = 9;
            // 
            // button_WriteFormula
            // 
            this.button_WriteFormula.Location = new System.Drawing.Point(295, 64);
            this.button_WriteFormula.Name = "button_WriteFormula";
            this.button_WriteFormula.Size = new System.Drawing.Size(133, 42);
            this.button_WriteFormula.TabIndex = 10;
            this.button_WriteFormula.Text = "Рассчитать по формуле";
            this.button_WriteFormula.UseVisualStyleBackColor = true;
            this.button_WriteFormula.Click += new System.EventHandler(this.button_WriteFormula_Click);
            // 
            // button_ReadCell
            // 
            this.button_ReadCell.Location = new System.Drawing.Point(295, 140);
            this.button_ReadCell.Name = "button_ReadCell";
            this.button_ReadCell.Size = new System.Drawing.Size(133, 42);
            this.button_ReadCell.TabIndex = 11;
            this.button_ReadCell.Text = "Считать данные";
            this.button_ReadCell.UseVisualStyleBackColor = true;
            this.button_ReadCell.Click += new System.EventHandler(this.button_ReadCell_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(463, 64);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 30);
            this.label5.TabIndex = 12;
            this.label5.Text = "Результат \r\nиз Excel";
            // 
            // textBox_ResultFormula
            // 
            this.textBox_ResultFormula.Location = new System.Drawing.Point(532, 71);
            this.textBox_ResultFormula.Name = "textBox_ResultFormula";
            this.textBox_ResultFormula.Size = new System.Drawing.Size(100, 23);
            this.textBox_ResultFormula.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(463, 147);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(60, 30);
            this.label6.TabIndex = 14;
            this.label6.Text = "Результат\r\nиз Excel";
            // 
            // textBox_ResultData
            // 
            this.textBox_ResultData.Location = new System.Drawing.Point(532, 147);
            this.textBox_ResultData.Name = "textBox_ResultData";
            this.textBox_ResultData.Size = new System.Drawing.Size(100, 23);
            this.textBox_ResultData.TabIndex = 15;
            // 
            // button_RunExcel
            // 
            this.button_RunExcel.Location = new System.Drawing.Point(12, 267);
            this.button_RunExcel.Name = "button_RunExcel";
            this.button_RunExcel.Size = new System.Drawing.Size(221, 39);
            this.button_RunExcel.TabIndex = 16;
            this.button_RunExcel.Text = "Запустить новое приложение Excel";
            this.button_RunExcel.UseVisualStyleBackColor = true;
            this.button_RunExcel.Click += new System.EventHandler(this.button_RunExcel_Click);
            // 
            // button_HideExcel
            // 
            this.button_HideExcel.Location = new System.Drawing.Point(239, 221);
            this.button_HideExcel.Name = "button_HideExcel";
            this.button_HideExcel.Size = new System.Drawing.Size(168, 39);
            this.button_HideExcel.TabIndex = 17;
            this.button_HideExcel.Text = "Скрыть Excel";
            this.button_HideExcel.UseVisualStyleBackColor = true;
            this.button_HideExcel.Click += new System.EventHandler(this.button_HideExcel_Click);
            // 
            // button_ShowExcel
            // 
            this.button_ShowExcel.Location = new System.Drawing.Point(239, 267);
            this.button_ShowExcel.Name = "button_ShowExcel";
            this.button_ShowExcel.Size = new System.Drawing.Size(168, 39);
            this.button_ShowExcel.TabIndex = 18;
            this.button_ShowExcel.Text = "Показать Excel";
            this.button_ShowExcel.UseVisualStyleBackColor = true;
            this.button_ShowExcel.Click += new System.EventHandler(this.button_ShowExcel_Click);
            // 
            // button_Close_Excel
            // 
            this.button_Close_Excel.Location = new System.Drawing.Point(414, 221);
            this.button_Close_Excel.Name = "button_Close_Excel";
            this.button_Close_Excel.Size = new System.Drawing.Size(223, 39);
            this.button_Close_Excel.TabIndex = 19;
            this.button_Close_Excel.Text = "Закрыть приложение Excel";
            this.button_Close_Excel.UseVisualStyleBackColor = true;
            this.button_Close_Excel.Click += new System.EventHandler(this.button_Close_Excel_Click);
            // 
            // button_Exit
            // 
            this.button_Exit.Location = new System.Drawing.Point(413, 267);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(224, 39);
            this.button_Exit.TabIndex = 20;
            this.button_Exit.Text = "Закрыть приложение";
            this.button_Exit.UseVisualStyleBackColor = true;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // form_Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(649, 322);
            this.Controls.Add(this.button_Exit);
            this.Controls.Add(this.button_Close_Excel);
            this.Controls.Add(this.button_ShowExcel);
            this.Controls.Add(this.button_HideExcel);
            this.Controls.Add(this.button_RunExcel);
            this.Controls.Add(this.textBox_ResultData);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox_ResultFormula);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button_ReadCell);
            this.Controls.Add(this.button_WriteFormula);
            this.Controls.Add(this.textBox_CellData);
            this.Controls.Add(this.textBox_CellFormula);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox_InputFormula);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button_WriteColumn);
            this.Controls.Add(this.listBox_InputData);
            this.Controls.Add(this.listBox_InputCell);
            this.Controls.Add(this.label1);
            this.Name = "form_Main";
            this.Text = "Программа для расчёта формул Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBox_InputCell;
        private System.Windows.Forms.ListBox listBox_InputData;
        private System.Windows.Forms.Button button_WriteColumn;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_InputFormula;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_CellFormula;
        private System.Windows.Forms.TextBox textBox_CellData;
        private System.Windows.Forms.Button button_WriteFormula;
        private System.Windows.Forms.Button button_ReadCell;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox_ResultFormula;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_ResultData;
        private System.Windows.Forms.Button button_RunExcel;
        private System.Windows.Forms.Button button_HideExcel;
        private System.Windows.Forms.Button button_ShowExcel;
        private System.Windows.Forms.Button button_Close_Excel;
        private System.Windows.Forms.Button button_Exit;
    }
}

