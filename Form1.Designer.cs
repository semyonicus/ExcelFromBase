using ExcelFromBase.Properties;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelFromBase
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.QueryArea = new System.Windows.Forms.TextBox();
            this.lvlServerName = new System.Windows.Forms.Label();
            this.StartWork = new System.Windows.Forms.Button();
            this.ServerName = new System.Windows.Forms.TextBox();
            this.StatusText = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ResultFile = new System.Windows.Forms.TextBox();
            this.DataBaseName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.PresetData = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.LoadSettings = new System.Windows.Forms.Button();
            this.listBoxConfigs = new System.Windows.Forms.ListBox();
            this.LoadChoiceToArea = new System.Windows.Forms.Button();
            this.DisableMessages = new System.Windows.Forms.CheckBox();
            this.UserNameTxt = new System.Windows.Forms.Label();
            this.TemplateFileName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.UseTemplate = new System.Windows.Forms.CheckBox();
            this.UploadExcel = new System.Windows.Forms.Button();
            this.LoginName = new System.Windows.Forms.TextBox();
            this.Login2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // QueryArea
            // 
            this.QueryArea.Location = new System.Drawing.Point(12, 304);
            this.QueryArea.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.QueryArea.Multiline = true;
            this.QueryArea.Name = "QueryArea";
            this.QueryArea.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.QueryArea.Size = new System.Drawing.Size(936, 72);
            this.QueryArea.TabIndex = 0;
            // 
            // lvlServerName
            // 
            this.lvlServerName.AutoSize = true;
            this.lvlServerName.Location = new System.Drawing.Point(28, 55);
            this.lvlServerName.Name = "lvlServerName";
            this.lvlServerName.Size = new System.Drawing.Size(75, 13);
            this.lvlServerName.TabIndex = 2;
            this.lvlServerName.Text = "Имя Сервера";
            this.lvlServerName.Visible = false;
            // 
            // StartWork
            // 
            this.StartWork.BackColor = System.Drawing.SystemColors.Highlight;
            this.StartWork.ForeColor = System.Drawing.SystemColors.Control;
            this.StartWork.Location = new System.Drawing.Point(244, 467);
            this.StartWork.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.StartWork.Name = "StartWork";
            this.StartWork.Size = new System.Drawing.Size(410, 33);
            this.StartWork.TabIndex = 0;
            this.StartWork.Text = "Шаг 3. Нажмите эту кнопку, и дождитесь результата";
            this.StartWork.UseVisualStyleBackColor = false;
            this.StartWork.Click += new System.EventHandler(this.StartWork_Click);
            // 
            // ServerName
            // 
            this.ServerName.Location = new System.Drawing.Point(31, 77);
            this.ServerName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ServerName.Name = "ServerName";
            this.ServerName.Size = new System.Drawing.Size(95, 20);
            this.ServerName.TabIndex = 1;
            this.ServerName.Text = "10.2.8.6";
            // 
            // StatusText
            // 
            this.StatusText.BackColor = System.Drawing.SystemColors.InfoText;
            this.StatusText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StatusText.ForeColor = System.Drawing.Color.LimeGreen;
            this.StatusText.Location = new System.Drawing.Point(49, 418);
            this.StatusText.Name = "StatusText";
            this.StatusText.Size = new System.Drawing.Size(898, 47);
            this.StatusText.TabIndex = 3;
            this.StatusText.Text = "Последовательно выполняйте шаги при открытии программы";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(50, 383);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(211, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Полный путь к имени файла результата";
            // 
            // ResultFile
            // 
            this.ResultFile.Location = new System.Drawing.Point(267, 380);
            this.ResultFile.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ResultFile.Name = "ResultFile";
            this.ResultFile.Size = new System.Drawing.Size(186, 20);
            this.ResultFile.TabIndex = 4;
            this.ResultFile.Text = "c:\\temp\\result.xlsx";
            // 
            // DataBaseName
            // 
            this.DataBaseName.Location = new System.Drawing.Point(590, 77);
            this.DataBaseName.Name = "DataBaseName";
            this.DataBaseName.Size = new System.Drawing.Size(86, 20);
            this.DataBaseName.TabIndex = 6;
            this.DataBaseName.Text = "YSUMain";
            this.DataBaseName.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(421, 55);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(454, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Имя базы данных где есть доступ, напишите Деканат, если нет доступа к Абитуриенты" +
    "";
            this.label3.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(49, 109);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(736, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Файл настроек json, от него зависит выгрузка если его нет вам придется вручную за" +
    "полнить окно с настройками, файл есть в сетевой папке";
            // 
            // PresetData
            // 
            this.PresetData.Location = new System.Drawing.Point(31, 124);
            this.PresetData.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.PresetData.Name = "PresetData";
            this.PresetData.Size = new System.Drawing.Size(186, 20);
            this.PresetData.TabIndex = 8;
            this.PresetData.Text = "c:\\temp\\config.json";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(332, 294);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(229, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Окно настроек вызовы хранимых процедур";
            // 
            // LoadSettings
            // 
            this.LoadSettings.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.LoadSettings.Location = new System.Drawing.Point(244, 124);
            this.LoadSettings.Name = "LoadSettings";
            this.LoadSettings.Size = new System.Drawing.Size(410, 23);
            this.LoadSettings.TabIndex = 11;
            this.LoadSettings.Text = "Шаг 1. Убедитесь что файл есть и Нажмите сюда";
            this.LoadSettings.UseVisualStyleBackColor = false;
            this.LoadSettings.Click += new System.EventHandler(this.LoadSettings_Click);
            // 
            // listBoxConfigs
            // 
            this.listBoxConfigs.FormattingEnabled = true;
            this.listBoxConfigs.Location = new System.Drawing.Point(12, 208);
            this.listBoxConfigs.Name = "listBoxConfigs";
            this.listBoxConfigs.Size = new System.Drawing.Size(936, 69);
            this.listBoxConfigs.TabIndex = 12;
            this.listBoxConfigs.SelectedIndexChanged += new System.EventHandler(this.ListBoxConfigs_SelectedIndexChanged);
            // 
            // LoadChoiceToArea
            // 
            this.LoadChoiceToArea.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.LoadChoiceToArea.Location = new System.Drawing.Point(244, 268);
            this.LoadChoiceToArea.Name = "LoadChoiceToArea";
            this.LoadChoiceToArea.Size = new System.Drawing.Size(410, 23);
            this.LoadChoiceToArea.TabIndex = 13;
            this.LoadChoiceToArea.Text = "Шаг 2. Выберите из списка нужную настройку и нажмите эту кнопку";
            this.LoadChoiceToArea.UseVisualStyleBackColor = false;
            this.LoadChoiceToArea.Click += new System.EventHandler(this.LoadChoiceToArea_Click);
            // 
            // DisableMessages
            // 
            this.DisableMessages.AutoSize = true;
            this.DisableMessages.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DisableMessages.ForeColor = System.Drawing.Color.Red;
            this.DisableMessages.Location = new System.Drawing.Point(541, 27);
            this.DisableMessages.Name = "DisableMessages";
            this.DisableMessages.Size = new System.Drawing.Size(406, 17);
            this.DisableMessages.TabIndex = 14;
            this.DisableMessages.Text = "Я все умею мне не нужны всплывающие подсказки  и мигание";
            this.DisableMessages.UseVisualStyleBackColor = true;
            this.DisableMessages.CheckedChanged += new System.EventHandler(this.Timer_Stop);
            // 
            // UserNameTxt
            // 
            this.UserNameTxt.AutoSize = true;
            this.UserNameTxt.Location = new System.Drawing.Point(28, 38);
            this.UserNameTxt.Name = "UserNameTxt";
            this.UserNameTxt.Size = new System.Drawing.Size(105, 13);
            this.UserNameTxt.TabIndex = 16;
            this.UserNameTxt.Text = "Имя Пользователя";
            this.UserNameTxt.Visible = false;
            // 
            // TemplateFileName
            // 
            this.TemplateFileName.Location = new System.Drawing.Point(31, 162);
            this.TemplateFileName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TemplateFileName.Name = "TemplateFileName";
            this.TemplateFileName.Size = new System.Drawing.Size(186, 20);
            this.TemplateFileName.TabIndex = 17;
            this.TemplateFileName.Text = "c:\\temp\\template.xlsx";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(223, 165);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 13);
            this.label6.TabIndex = 18;
            // 
            // UseTemplate
            // 
            this.UseTemplate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.UseTemplate.ForeColor = System.Drawing.Color.MidnightBlue;
            this.UseTemplate.Location = new System.Drawing.Point(244, 148);
            this.UseTemplate.Name = "UseTemplate";
            this.UseTemplate.Size = new System.Drawing.Size(551, 48);
            this.UseTemplate.TabIndex = 19;
            this.UseTemplate.Text = "Использовать данный шаблон, если стоит галочка уже будет все создаваться на основ" +
    "ании этого файла, нужно если  у вас есть готовый шаблон, например при создании р" +
    "асписания есть готовый шаблон";
            this.UseTemplate.UseVisualStyleBackColor = true;
            this.UseTemplate.CheckedChanged += new System.EventHandler(this.UseTemplate_CheckedChanged);
            // 
            // UploadExcel
            // 
            this.UploadExcel.Location = new System.Drawing.Point(12, 12);
            this.UploadExcel.Name = "UploadExcel";
            this.UploadExcel.Size = new System.Drawing.Size(211, 23);
            this.UploadExcel.TabIndex = 25;
            this.UploadExcel.Text = "Перейти к загрузке Excel в базу";
            this.UploadExcel.UseVisualStyleBackColor = true;
            this.UploadExcel.Click += new System.EventHandler(this.UploadExcel_Click);
            // 
            // LoginName
            // 
            this.LoginName.Location = new System.Drawing.Point(304, 38);
            this.LoginName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.LoginName.Name = "LoginName";
            this.LoginName.Size = new System.Drawing.Size(95, 20);
            this.LoginName.TabIndex = 20;
            this.LoginName.Text = "user";
            // 
            // Login2
            // 
            this.Login2.Location = new System.Drawing.Point(304, 77);
            this.Login2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Login2.Name = "Login2";
            this.Login2.Size = new System.Drawing.Size(95, 20);
            this.Login2.TabIndex = 21;
            this.Login2.Text = "1";
            this.Login2.UseSystemPasswordChar = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(172, 41);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(126, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "ИИСУСС пользователь";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(172, 77);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(91, 13);
            this.label9.TabIndex = 24;
            this.label9.Text = "ИИСУСС пароль";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(959, 591);
            this.Controls.Add(this.UploadExcel);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.Login2);
            this.Controls.Add(this.LoginName);
            this.Controls.Add(this.UseTemplate);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.TemplateFileName);
            this.Controls.Add(this.UserNameTxt);
            this.Controls.Add(this.DisableMessages);
            this.Controls.Add(this.LoadChoiceToArea);
            this.Controls.Add(this.listBoxConfigs);
            this.Controls.Add(this.LoadSettings);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.PresetData);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.DataBaseName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ResultFile);
            this.Controls.Add(this.StatusText);
            this.Controls.Add(this.lvlServerName);
            this.Controls.Add(this.ServerName);
            this.Controls.Add(this.StartWork);
            this.Controls.Add(this.QueryArea);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Копирование Результата хранимой процедуры в Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox QueryArea;
        private System.Windows.Forms.TextBox ServerName;
        private System.Windows.Forms.TextBox ResultFile;
        private System.Windows.Forms.TextBox DataBaseName;
        private System.Windows.Forms.Label lvlServerName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label StatusText;
        private Button StartWork;
        private Label label4;
        private TextBox PresetData;
        private Label label5;
        private Button LoadSettings;
        private ListBox listBoxConfigs;
        private Button LoadChoiceToArea;
        private CheckBox DisableMessages;
        private Label UserNameTxt;
        private TextBox TemplateFileName;
        private Label label6;
        private CheckBox UseTemplate;
        private Button UploadExcel;
        private TextBox LoginName;
        private TextBox Login2;
        private Label label7;
        private Label label9;
    }
}

