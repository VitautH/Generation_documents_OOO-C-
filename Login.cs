using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
namespace OnCore
{
    class Login
    {
        private static string password_hash = "249d4e93c29a24956d900c55fd162822";
        private static string login = "admin";
    
        private bool is_login;

        public Login(string login, string password)
        {
            Autufication(login, password);


        }
    

        private bool Autufication(string login, string password)
        {
            if (login == Login.login)
            {
                string source = password;
                using (MD5 md5Hash = MD5.Create())
                {
                    bool request = Main.VerifyMd5Hash(md5Hash, password, password_hash);
                    if (request)
                    {

                        is_login = true;
                        return true;

                    }
                    else
                    {

                        is_login = false;
                        return false;
                    }

                }
            }
            else
            {
                return false;
            }


        
           

        }
        

        public bool Is_login()
        {
            return is_login;
        }
    }

}
