using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using  System.Threading;
using System.Security.Cryptography;

namespace OnCore
{

    public partial class Main_form : Form
    {

      
        public Main_form()
        {
            InitializeComponent();
        }

      


        private void Main_form_Load(object sender, EventArgs e)
        {

            groupBox_setting.Visible = false;
            folder_text.ReadOnly = true;
            Main nt = new Main();
            folder_text.Text = nt.Get_register_key("folder");
            string serial_number = nt.Get_serial_key();
            string key = nt.Get_register_key("secret_key");
            using (MD5 md5Hash = MD5.Create())
            {

             bool request =   Main.VerifyMd5Hash(md5Hash, serial_number, key);
                if (!request)
                {
                    MessageBox.Show("Программа защищена от повторного использования!");
                    Application.Exit();
                }
            }
           

            nt = null;

        }

        private void Login_Click(object sender, EventArgs e)
        {
            Login login_obj = new Login (loginBox.Text, passwordBox.Text);


            if (login_obj.Is_login())
            {
              
                groupBox_login.Visible = false;
                groupBox_setting.Visible = true;

            }
            else
            {
                MessageBox.Show("Не верное имя пользователя или пароль!");
            }

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void loginBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter_1(object sender, EventArgs e)
        {

        }

        private void label1_Click_2(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
           Main nt = new Main();
           
            if (nt.Save_register_key("folder", folder_text.Text))
            {

            }
            else
            {
                MessageBox.Show("Ошибка записи в  реестр!");
            }
            nt = null;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (folder_text.Text == "")
            {
                MessageBox.Show("Выберите папку для сохранения документов!");
            }
          else
            {
                if (select_generation_document.SelectedItem == "Документы ООО")
                {

                    bool is_show_forms = false;

                    foreach (Form f in Application.OpenForms)
                    {
                        is_show_forms = f.Name == "Form_generation_ooo";
                        if (is_show_forms)
                        {
                            f.Activate();
                            break;
                        }
                    }
                    if (!is_show_forms)
                    {
                        Form_generation_ooo nt = new Form_generation_ooo();

                        nt.Show();



                    }




                }
                else
                {
                    MessageBox.Show("Выберите тип документа!");
                }
            }

        }

        private void button1_Click_2(object sender, EventArgs e)
        {

        }

        private void folder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.ShowNewFolderButton = false;
            if (FBD.ShowDialog() == DialogResult.OK)
            {
               folder_text.Text =FBD.SelectedPath;
               
            }
        }

        private void folder_text_TextChanged(object sender, EventArgs e)
        {
          


        }
    }
}
