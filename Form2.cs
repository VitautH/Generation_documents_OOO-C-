using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;


namespace OnCore
{
    public partial class Form_generation_ooo : Form
    {
        

        public Form_generation_ooo()
        {
            InitializeComponent();
            GC.Collect();


        }


        private void Form_generation_ooo_Load(object sender, EventArgs e)
        {

           
            tabPage5.Parent = null;
            tabPage6.Parent = null;
            city_ur.Checked = true;
            city_director.Checked = true;
            passport_director.Checked = true;
            city_founder_1.Checked = true;
            city_founder_2.Checked = true;
            city_founder_3.Checked = true;

            founder_1_sex_male.Checked = true;
            founder_2_sex_male.Checked = true;
            founder_3_sex_male.Checked = true;
            founder_1_passport.Checked = true;
            founder_2_passport.Checked = true;
            founder_3_passport.Checked = true;
            progressBar.Visible = false;
            label_run.Visible = false;
           if (quantity_founder.Value == 1 ) { founder_1_precent_fond.Text = "100";}
        }



        private void send_Click(object sender, EventArgs e)
        {



            IList<bool> validation = new List<bool>();
            IDictionary<string, string> generation_ooo_dictionary = new Dictionary<string, string>();

            foreach (Control ctrl in this.tabPage1.Controls)
            {
                if (ctrl is DateTimePicker)
                {
                    string key = ((DateTimePicker)ctrl).Name;
                    string value = ((DateTimePicker)ctrl).Value.ToString("dd.MM.yyyy");

                    generation_ooo_dictionary.Add(key, value);

                }
                if (ctrl is TextBox)
                {

                    if (((TextBox) ctrl).Text == "")
                    {
                        validation.Add(false);
                        ((TextBox) ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                       
                        ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((TextBox) ctrl).Name;
                        string value = ((TextBox) ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
                if (ctrl is RichTextBox)
                {
                    if (((RichTextBox) ctrl).Text == "")
                    {
                        validation.Add(false);
                        ((RichTextBox) ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                        
                        ((RichTextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((RichTextBox) ctrl).Name;
                        string value = ((RichTextBox) ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }

                if (ctrl is MaskedTextBox)
                {
                    if (!((MaskedTextBox) ctrl).MaskFull)
                    {
                        validation.Add(false);
                        ((MaskedTextBox) ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                   
                        ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((MaskedTextBox) ctrl).Name;
                        string value = ((MaskedTextBox) ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
           
            }

            foreach (Control ctrl in this.tabPage2.Controls)
            {
                if (ctrl is TextBox)
                {
                    if (((TextBox) ctrl).Text == "")
                    {
                        validation.Add(false);
                        ((TextBox) ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                        ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((TextBox) ctrl).Name;
                        string value = ((TextBox) ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
                if (ctrl is MaskedTextBox)
                {
                    if (!((MaskedTextBox) ctrl).MaskFull)
                    {
                        validation.Add(false);
                        ((MaskedTextBox) ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                        ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((MaskedTextBox) ctrl).Name;
                        string value = ((MaskedTextBox) ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
             
            }



            foreach (Control ctrl in this.tabPage3.Controls)
            {

                if (ctrl is DateTimePicker)
                {
                    string key = ((DateTimePicker)ctrl).Name;
                    string value = ((DateTimePicker)ctrl).Value.ToString("dd.MM.yyyy");

                    generation_ooo_dictionary.Add(key, value);

                }
                if (ctrl is TextBox)
                {
                    if (((TextBox)ctrl).Text == "")
                    {
                        validation.Add(false);
                        ((TextBox)ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                        ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((TextBox)ctrl).Name;
                        string value = ((TextBox)ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
                if (ctrl is MaskedTextBox)
                {
                    if (!((MaskedTextBox)ctrl).MaskFull)
                    {
                        validation.Add(false);
                        ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                        ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((MaskedTextBox)ctrl).Name;
                        string value = ((MaskedTextBox)ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
               
            }



            foreach (Control ctrl in this.tabPage4.Controls)
            {

                if (ctrl is DateTimePicker)
                {
                    string key = ((DateTimePicker)ctrl).Name;
                    string value = ((DateTimePicker)ctrl).Value.ToString("dd.MM.yyyy");

                    generation_ooo_dictionary.Add(key, value);

                }
                if (ctrl is TextBox)
                {
                    if (((TextBox)ctrl).Text == "")
                    {
                        validation.Add(false);
                        ((TextBox)ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                       
                        ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((TextBox)ctrl).Name;
                        string value = ((TextBox)ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
                if (ctrl is MaskedTextBox)
                {
                    if (!((MaskedTextBox)ctrl).MaskFull)
                    {
                        validation.Add(false);
                        ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.Red;

                    }
                    else
                    {
                        
                        ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                        string key = ((MaskedTextBox)ctrl).Name;
                        string value = ((MaskedTextBox)ctrl).Text;

                        generation_ooo_dictionary.Add(key, value);
                    }
                }
                 
            }



            if (quantity_founder.Value == 2)
            {

                foreach (Control ctrl in this.tabPage5.Controls)
                {

                    if (ctrl is DateTimePicker)
                    {
                        string key = ((DateTimePicker)ctrl).Name;
                        string value = ((DateTimePicker)ctrl).Value.ToString("dd.MM.yyyy");

                        generation_ooo_dictionary.Add(key, value);

                    }
                    if (ctrl is TextBox)
                    {
                        if (((TextBox) ctrl).Text == "")
                        {
                            validation.Add(false);
                            ((TextBox) ctrl).BackColor = System.Drawing.Color.Red;

                        }
                        else
                        {
                            ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                            string key = ((TextBox) ctrl).Name;
                            string value = ((TextBox) ctrl).Text;

                            generation_ooo_dictionary.Add(key, value);
                        }
                    }
                    if (ctrl is MaskedTextBox)
                    {
                        if (!((MaskedTextBox) ctrl).MaskFull)
                        {
                            validation.Add(false);
                            ((MaskedTextBox) ctrl).BackColor = System.Drawing.Color.Red;

                        }
                        else
                        {
                            ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                            string key = ((MaskedTextBox) ctrl).Name;
                            string value = ((MaskedTextBox) ctrl).Text;

                            generation_ooo_dictionary.Add(key, value);
                        }
                    }
                }

               
            }
            if (quantity_founder.Value == 3)
            {

                foreach (Control ctrl in this.tabPage5.Controls)
                {

                    if (ctrl is DateTimePicker)
                    {
                        string key = ((DateTimePicker)ctrl).Name;
                        string value = ((DateTimePicker)ctrl).Value.ToString("dd.MM.yyyy");

                        generation_ooo_dictionary.Add(key, value);

                    }
                    if (ctrl is TextBox)
                    {
                        if (((TextBox) ctrl).Text == "")
                        {
                            validation.Add(false);
                            ((TextBox) ctrl).BackColor = System.Drawing.Color.Red;

                        }
                        else
                        {
                            ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                            string key = ((TextBox) ctrl).Name;
                            string value = ((TextBox) ctrl).Text;

                            generation_ooo_dictionary.Add(key, value);
                        }
                    }
                    if (ctrl is MaskedTextBox)
                    {
                        if (!((MaskedTextBox) ctrl).MaskFull)
                        {
                            validation.Add(false);
                            ((MaskedTextBox) ctrl).BackColor = System.Drawing.Color.Red;

                        }
                        else
                        {
                            ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                            string key = ((MaskedTextBox) ctrl).Name;
                            string value = ((MaskedTextBox) ctrl).Text;

                            generation_ooo_dictionary.Add(key, value);
                        }
                    }
                }

                foreach (Control ctrl in this.tabPage6.Controls)
                {

                    if (ctrl is DateTimePicker)
                    {
                        string key = ((DateTimePicker)ctrl).Name;
                        string value = ((DateTimePicker)ctrl).Value.ToString("dd.MM.yyyy");

                        generation_ooo_dictionary.Add(key, value);

                    }
                    if (ctrl is TextBox)
                    {
                        if (((TextBox)ctrl).Text == "")
                        {
                            validation.Add(false);
                            ((TextBox)ctrl).BackColor = System.Drawing.Color.Red;

                        }
                        else
                        {
                            ((TextBox)ctrl).BackColor = System.Drawing.Color.White;
                            string key = ((TextBox)ctrl).Name;
                            string value = ((TextBox)ctrl).Text;

                            generation_ooo_dictionary.Add(key, value);
                        }
                    }
                    if (ctrl is MaskedTextBox)
                    {
                        if (!((MaskedTextBox)ctrl).MaskFull)
                        {
                            validation.Add(false);
                            ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.Red;

                        }
                        else
                        {
                            ((MaskedTextBox)ctrl).BackColor = System.Drawing.Color.White;
                            string key = ((MaskedTextBox)ctrl).Name;
                            string value = ((MaskedTextBox)ctrl).Text;

                            generation_ooo_dictionary.Add(key, value);
                        }
                    }
                }

               
            }
            ////// Radio Buuton add in Collection
            foreach (Control control in this.locality_ur.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {


                        generation_ooo_dictionary.Add("locality_ur", radio.Name);
                    }
                }

            }
            foreach (Control control in this.locality_director.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {


                        generation_ooo_dictionary.Add("locality_director", radio.Name);
                    }
                }
            }
            foreach (Control control in this.director_type_document.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {


                        generation_ooo_dictionary.Add("director_type_document", radio.Name);
                    }
                }
            }
            foreach (Control control in this.founder_1_sex.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {


                        generation_ooo_dictionary.Add("founder_1_sex", radio.Name);
                    }
                }
            }
            foreach (Control control in this.locality_founder_1.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {


                        generation_ooo_dictionary.Add("locality_founder_1", radio.Name);
                    }
                }
            }
            foreach (Control control in this.founder_1_type_document.Controls)
            {
                if (control is RadioButton)
                {
                    RadioButton radio = control as RadioButton;
                    if (radio.Checked)
                    {


                        generation_ooo_dictionary.Add("founder_1_type_document", radio.Name);
                    }
                }
            }
            if (quantity_founder.Value == 2)
            {
                foreach (Control control in this.founder_2_sex.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("founder_2_sex", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.locality_founder_2.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("locality_founder_2", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.founder_2_type_document.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("founder_2_type_document", radio.Name);
                        }
                    }
                }
            }
            if (quantity_founder.Value == 3)
            {
                foreach (Control control in this.founder_2_sex.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("founder_2_sex", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.locality_founder_2.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("locality_founder_2", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.founder_2_type_document.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("founder_2_type_document", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.founder_3_sex.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("founder_3_sex", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.locality_founder_3.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("locality_founder_3", radio.Name);
                        }
                    }
                }
                foreach (Control control in this.founder_3_type_document.Controls)
                {
                    if (control is RadioButton)
                    {
                        RadioButton radio = control as RadioButton;
                        if (radio.Checked)
                        {


                            generation_ooo_dictionary.Add("founder_3_type_document", radio.Name);
                        }
                    }
                }
            }
            generation_ooo_dictionary.Add("quantity_founder", quantity_founder.Value.ToString());
         

                               if (validation.Count == 0)
                               {
                                      
                    validation.Clear();
                                   send.Visible = false;
                progressBar.Visible = true;
                                   label_run.Visible = true;
           
                 // Конструктор
                 Generation_OOO_docx docx = new Generation_OOO_docx(generation_ooo_dictionary);
                //// Запуск потока //
                Thread myThread = new Thread(new ThreadStart(docx.Generate)); //Создаем новый объект потока (Thread)
                myThread.IsBackground=false;
                myThread.Name = "Generation_OOO_docx";

                myThread.Start();
                                   progressBar.Value = 0;
                                
                                   while ( myThread.IsAlive)
                                   {

                                       if (progressBar.Value < 1000)
                                       {
                                           progressBar.Value++;


                                       }
                                       else
                                       {
                                          
                        send.Visible = true;
                        progressBar.Visible = false;
                                           label_run.Visible = false;

                        break;
                                       }
                 


                }
                         

                                 


                               }
                else
               {
                                    
                                        MessageBox.Show("Вы заполнили не все поля формы!");
                    
                                       generation_ooo_dictionary.Clear();
                   generation_ooo_dictionary = null;
                                      validation.Clear();
                   validation = null;
                   GC.Collect();
               }
            
           
               
            


        }

      
        




        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void next_to_tabpage1_Click(object sender, EventArgs e)
        {
          
                tabControl1.SelectedIndex++;
            
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

           
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(all_fond.Text, "  ^ [0-9]"))
            {
                all_fond.Text = "";
            }
            if (quantity_founder.Value == 1)
            {
                founder_1_fond.Text = all_fond.Text;
            }
            else
            {
                founder_1_fond.Text = "";
            }
        }

        private void all_fond_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (quantity_founder.Value == 1)
            {
                founder_1_fond.Text = all_fond.Text;
            }
            else
            {
                founder_1_fond.Text = "";
            }

            if (quantity_founder.Value == 1 )
            {
                founder_1_precent_fond.Text = "100";
            }
            else
            {
                founder_1_precent_fond.Text = "";
            }
            finish.Parent = null;
            if (quantity_founder.Value == 1)
            {
                tabPage5.Parent = null;
                tabPage6.Parent = null;
                finish.Parent = tabControl1;

            }
            if (quantity_founder.Value == 2)
            {

                tabPage6.Parent = null;
                tabPage5.Parent = tabControl1;
                finish.Parent = tabControl1;
            }
            if (quantity_founder.Value == 3)
            {
                tabPage5.Parent = tabControl1;
                tabPage6.Parent = tabControl1;
                finish.Parent = tabControl1;
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
           
        }

      

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

      

        private void button2_Click(object sender, EventArgs e)
        {
            
                tabControl1.SelectedIndex--;
            
        }

       

        private void registering_authority_TextChanged(object sender, EventArgs e)
        {
           

        }

     

        private void type_city_ur_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
         
          
            tabControl1.SelectedIndex++;
            
        }


        private void soglas_TextChanged(object sender, EventArgs e)
        {
            

        }

        private void oked_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(oked.Text, "  ^ [0-9]"))
            {
                oked.Text = "";
            }

        }

        private void oked_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }




        private void index_ur_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(index_ur.Text, "  ^ [0-9]"))
            {
               index_ur.Text = "";
            }


        }
        private void index_ur_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void region_ur_TextChanged(object sender, EventArgs e)
        {

          
        }

        private void district_ur_TextChanged(object sender, EventArgs e)
        {

           
        }

        private void city_ur_TextChanged(object sender, EventArgs e)
        {

          
        }

        private void street_ur_TextChanged(object sender, EventArgs e)
        {

         
        }

        private void house_ur_TextChanged(object sender, EventArgs e)
        {
           
        }

       

        private void textBox71_TextChanged(object sender, EventArgs e)
        {

        }

        private void prev_to_tabpage1_Click(object sender, EventArgs e)
        {
           
                tabControl1.SelectedIndex--;
            
        }

        private void next_to_tabpage3_Click(object sender, EventArgs e)
        {
            
                tabControl1.SelectedIndex++;
            
        }

        private void prev_to_tabpage3_Click(object sender, EventArgs e)
        {
            
                tabControl1.SelectedIndex--;
            

        }

        private void button8_Click(object sender, EventArgs e)
        {
            
                tabControl1.SelectedIndex = 6;
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
                tabControl1.SelectedIndex = 4;
            

        }

        private void button4_Click(object sender, EventArgs e)
        {
             
                    tabControl1.SelectedIndex = 4;
                

            

        }

        private void button7_Click(object sender, EventArgs e)
        {
           

                tabControl1.SelectedIndex = 3;


            
        }

        private void button6_Click(object sender, EventArgs e)
        {
          

                tabControl1.SelectedIndex = 5;


            
        }

        private void finish_Click(object sender, EventArgs e)
        {

        }

        private void progress_bar_Click(object sender, EventArgs e)
        {

        }

        private void founder_1_fond_TextChanged(object sender, EventArgs e)
        {

        }

        private void index_director_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(index_director.Text, "  ^ [0-9]"))
            {
                index_director.Text = "";
            }
        }
        private void index_director_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void phone_director_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(phone_director.Text, "  ^ [0-9]"))
            {
               phone_director.Text = "";
            }

        }
        private void phone_director_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void founder_1_index_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(phone_director.Text, "  ^ [0-9]"))
            {
                founder_1_index.Text = "";
            }
        }
        private void founder_1_index_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void label174_Click(object sender, EventArgs e)
        {

        }
    }
}
