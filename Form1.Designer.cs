namespace OnCore
{
    partial class Main_form
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
            this.Login = new System.Windows.Forms.Button();
            this.passwordBox = new System.Windows.Forms.TextBox();
            this.loginBox = new System.Windows.Forms.TextBox();
            this.groupBox_login = new System.Windows.Forms.GroupBox();
            this.label_password = new System.Windows.Forms.Label();
            this.label_login = new System.Windows.Forms.Label();
            this.label_login_title = new System.Windows.Forms.Label();
            this.groupBox_setting = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.folder_text = new System.Windows.Forms.TextBox();
            this.folder = new System.Windows.Forms.Button();
            this.button_next = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.select_generation_document = new System.Windows.Forms.ComboBox();
            this.save_setting = new System.Windows.Forms.Button();
            this.label_setting = new System.Windows.Forms.Label();
            this.groupBox_login.SuspendLayout();
            this.groupBox_setting.SuspendLayout();
            this.SuspendLayout();
            // 
            // Login
            // 
            this.Login.Location = new System.Drawing.Point(118, 126);
            this.Login.Name = "Login";
            this.Login.Size = new System.Drawing.Size(133, 22);
            this.Login.TabIndex = 1;
            this.Login.Text = "Войти";
            this.Login.Click += new System.EventHandler(this.Login_Click);
            // 
            // passwordBox
            // 
            this.passwordBox.Location = new System.Drawing.Point(118, 86);
            this.passwordBox.MaxLength = 14;
            this.passwordBox.Name = "passwordBox";
            this.passwordBox.PasswordChar = '*';
            this.passwordBox.Size = new System.Drawing.Size(235, 20);
            this.passwordBox.TabIndex = 0;
            // 
            // loginBox
            // 
            this.loginBox.Location = new System.Drawing.Point(118, 49);
            this.loginBox.Name = "loginBox";
            this.loginBox.Size = new System.Drawing.Size(235, 20);
            this.loginBox.TabIndex = 0;
            this.loginBox.TextChanged += new System.EventHandler(this.loginBox_TextChanged);
            // 
            // groupBox_login
            // 
            this.groupBox_login.Controls.Add(this.label_password);
            this.groupBox_login.Controls.Add(this.label_login);
            this.groupBox_login.Controls.Add(this.label_login_title);
            this.groupBox_login.Controls.Add(this.Login);
            this.groupBox_login.Controls.Add(this.passwordBox);
            this.groupBox_login.Controls.Add(this.loginBox);
            this.groupBox_login.Location = new System.Drawing.Point(3, 3);
            this.groupBox_login.Name = "groupBox_login";
            this.groupBox_login.Size = new System.Drawing.Size(393, 204);
            this.groupBox_login.TabIndex = 2;
            this.groupBox_login.TabStop = false;
            this.groupBox_login.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label_password
            // 
            this.label_password.AutoSize = true;
            this.label_password.Location = new System.Drawing.Point(28, 89);
            this.label_password.Name = "label_password";
            this.label_password.Size = new System.Drawing.Size(45, 13);
            this.label_password.TabIndex = 4;
            this.label_password.Text = "Пароль";
            // 
            // label_login
            // 
            this.label_login.AutoSize = true;
            this.label_login.Location = new System.Drawing.Point(9, 52);
            this.label_login.Name = "label_login";
            this.label_login.Size = new System.Drawing.Size(103, 13);
            this.label_login.TabIndex = 3;
            this.label_login.Text = "Имя пользователя";
            this.label_login.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // label_login_title
            // 
            this.label_login_title.AutoSize = true;
            this.label_login_title.Font = new System.Drawing.Font("Arial Rounded MT Bold", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_login_title.Location = new System.Drawing.Point(178, 6);
            this.label_login_title.Name = "label_login_title";
            this.label_login_title.Size = new System.Drawing.Size(73, 26);
            this.label_login_title.TabIndex = 2;
            this.label_login_title.Text = "Войти";
            this.label_login_title.Click += new System.EventHandler(this.label1_Click);
            // 
            // groupBox_setting
            // 
            this.groupBox_setting.Controls.Add(this.label3);
            this.groupBox_setting.Controls.Add(this.folder_text);
            this.groupBox_setting.Controls.Add(this.folder);
            this.groupBox_setting.Controls.Add(this.button_next);
            this.groupBox_setting.Controls.Add(this.label2);
            this.groupBox_setting.Controls.Add(this.select_generation_document);
            this.groupBox_setting.Controls.Add(this.save_setting);
            this.groupBox_setting.Controls.Add(this.label_setting);
            this.groupBox_setting.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupBox_setting.Location = new System.Drawing.Point(0, 0);
            this.groupBox_setting.Name = "groupBox_setting";
            this.groupBox_setting.Size = new System.Drawing.Size(587, 297);
            this.groupBox_setting.TabIndex = 3;
            this.groupBox_setting.TabStop = false;
            this.groupBox_setting.Enter += new System.EventHandler(this.groupBox1_Enter_1);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Roboto Condensed", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(188, 17);
            this.label3.TabIndex = 9;
            this.label3.Text = "Папка для сохранения документов";
            // 
            // folder_text
            // 
            this.folder_text.Location = new System.Drawing.Point(249, 94);
            this.folder_text.Name = "folder_text";
            this.folder_text.Size = new System.Drawing.Size(206, 20);
            this.folder_text.TabIndex = 8;
            this.folder_text.TextChanged += new System.EventHandler(this.folder_text_TextChanged);
            // 
            // folder
            // 
            this.folder.Location = new System.Drawing.Point(471, 91);
            this.folder.Name = "folder";
            this.folder.Size = new System.Drawing.Size(75, 23);
            this.folder.TabIndex = 7;
            this.folder.Text = "Выбрать";
            this.folder.UseVisualStyleBackColor = true;
            this.folder.Click += new System.EventHandler(this.folder_Click);
            // 
            // button_next
            // 
            this.button_next.Location = new System.Drawing.Point(298, 224);
            this.button_next.Name = "button_next";
            this.button_next.Size = new System.Drawing.Size(75, 23);
            this.button_next.TabIndex = 6;
            this.button_next.Text = "Перейти";
            this.button_next.UseVisualStyleBackColor = true;
            this.button_next.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Roboto Condensed", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(15, 224);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Выберите тип документа";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // select_generation_document
            // 
            this.select_generation_document.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.select_generation_document.FormattingEnabled = true;
            this.select_generation_document.Items.AddRange(new object[] {
            "Документы ООО"});
            this.select_generation_document.Location = new System.Drawing.Point(158, 224);
            this.select_generation_document.Name = "select_generation_document";
            this.select_generation_document.Size = new System.Drawing.Size(121, 21);
            this.select_generation_document.TabIndex = 4;
            this.select_generation_document.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // save_setting
            // 
            this.save_setting.Location = new System.Drawing.Point(253, 128);
            this.save_setting.Name = "save_setting";
            this.save_setting.Size = new System.Drawing.Size(125, 23);
            this.save_setting.TabIndex = 3;
            this.save_setting.Text = "Сохранить";
            this.save_setting.UseVisualStyleBackColor = true;
            this.save_setting.Click += new System.EventHandler(this.button1_Click);
            // 
            // label_setting
            // 
            this.label_setting.AutoSize = true;
            this.label_setting.Font = new System.Drawing.Font("Lucida Sans", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_setting.Location = new System.Drawing.Point(186, 23);
            this.label_setting.Name = "label_setting";
            this.label_setting.Size = new System.Drawing.Size(192, 22);
            this.label_setting.TabIndex = 0;
            this.label_setting.Text = "OnCore Documents";
            this.label_setting.Click += new System.EventHandler(this.label1_Click_2);
            // 
            // Main_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(614, 297);
            this.Controls.Add(this.groupBox_setting);
            this.Controls.Add(this.groupBox_login);
            this.Name = "Main_form";
            this.Text = "OnCore Documents";
            this.Load += new System.EventHandler(this.Main_form_Load);
            this.groupBox_login.ResumeLayout(false);
            this.groupBox_login.PerformLayout();
            this.groupBox_setting.ResumeLayout(false);
            this.groupBox_setting.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

       
        private System.Windows.Forms.TextBox passwordBox;
        private System.Windows.Forms.TextBox loginBox;
        private System.Windows.Forms.Button Login;
        private System.Windows.Forms.GroupBox groupBox_login;
        private System.Windows.Forms.Label label_login_title;
        private System.Windows.Forms.Label label_password;
        private System.Windows.Forms.Label label_login;
        private System.Windows.Forms.GroupBox groupBox_setting;
        private System.Windows.Forms.Label label_setting;
        private System.Windows.Forms.Button save_setting;
        private System.Windows.Forms.ComboBox select_generation_document;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_next;
        private System.Windows.Forms.Button folder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox folder_text;
    }
}

