using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Ionic.Zip;

namespace OnCore
{
    class Docx: IDisposable, IDocx
    {

        private bool disposed; 
        private string zipPath;
        private string path_output;
        private string extractPath; // Папка tmp создаётся автоматически
        private string finaldocx;

       private ZipFile zf = null;
        private FileInfo f = null;
        public  Docx(string docx)
        {
            disposed = false;
            extractPath = Directory.GetCurrentDirectory() + @"\temp\";
          
            Main nt = new Main();
            path_output = nt.Get_register_key("folder");
            nt = null;
            try
            {
                zipPath = Directory.GetCurrentDirectory() + @"\template\" + docx;
                zf = new ZipFile(zipPath);
                zf.ExtractAll(extractPath);
             
            }
            catch
            {
                MessageBox.Show("Произошла ошибка! Возможно, файл "+docx+" повреждён или он не существует");
            }
       

        }
        public void SetValue(string key, string value)
        {
            try
            {
                StringBuilder sbText = new StringBuilder(File.ReadAllText(@"" + extractPath + @"\word\document.xml"));
                sbText.Replace("${" + key + "}", value);

                File.WriteAllText(@"" + extractPath + @"\word\document.xml", sbText.ToString());

                System.Console.ReadLine();
            }
            catch
            {
                MessageBox.Show("Произошла ошибка генерации документа!");
            }
        }
        public bool SaveDocument (string new_document)
        {

            try
            {

                zf = new ZipFile(new_document);
                zf.AddDirectory(@"" + extractPath);
                zf.Save();

            
            }
            catch
            {
                MessageBox.Show("Произошла ошибка создания документа "+ new_document+" !");
            }
            try
            {
                Directory.Delete(extractPath, true);
                f = new FileInfo(Directory.GetCurrentDirectory() + @"\" + new_document);
                if (File.Exists(path_output + @"\" + new_document))
                {
                    f.MoveTo(path_output + @"\" + DateTime.Today + "_" + new_document);

                }
                else
                {
                    f.MoveTo(path_output + @"\" + new_document);
                }
               
            }
            catch
            {
                MessageBox.Show("Произошла ошибка перемещения документа " + new_document + " в папку !");
            }
            return true;
            
        }

        public void Dispose()
        {
            lock (this)
            {
                if (!this.disposed)
                {
                   zf.Dispose();
                    f = null;
                   
                  
                }
                this.disposed = true;
                GC.SuppressFinalize(this);
              

            }
        }

        ~Docx()
        {
            this.Dispose();
        }


    }
}



 