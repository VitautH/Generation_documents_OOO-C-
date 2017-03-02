using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Cryptography;
using Microsoft.Win32;
using System.Management;


namespace OnCore
{
     class Main
     {


         private string request;

        public bool Save_register_key(string key, string value)
        {
            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey oncore_key = currentUserKey.CreateSubKey("OnCore");

            oncore_key.SetValue(key, value);
            oncore_key.Close();
            return true;
        }

        public string Get_register_key(string key)
        {
            try
            {
                RegistryKey currentUserKey = Registry.CurrentUser;

                RegistryKey oncore_key = currentUserKey.CreateSubKey("OnCore");

               request = (string) oncore_key.GetValue(key);
                oncore_key.Close();
                return request;
            }
            catch
            {
                MessageBox.Show("Произошла ошибка чтения из памяти!");
                Application.Exit();
            }
            return request;

        }

        public  string Get_serial_key()
        {
            List<string> ids =
                new List<string>();

            ManagementObjectSearcher searcher;

            //процессор
            searcher = new ManagementObjectSearcher("root\\CIMV2",
                "SELECT * FROM Win32_Processor");
            foreach (ManagementObject queryObj in searcher.Get())
                ids.Add(queryObj["ProcessorId"].ToString());

            //мать
            searcher = new ManagementObjectSearcher("root\\CIMV2",
                "SELECT * FROM CIM_Card");
            foreach (ManagementObject queryObj in searcher.Get())
                ids.Add(queryObj["SerialNumber"].ToString());



            //UUID
            searcher = new ManagementObjectSearcher("root\\CIMV2",
                "SELECT UUID FROM Win32_ComputerSystemProduct");
            foreach (ManagementObject queryObj in searcher.Get())
                ids.Add(queryObj["UUID"].ToString());
            string serial_key = "NT";
            foreach (var x in ids)
            {
                serial_key += x;





            }
            return serial_key;
        }

        public static string GetMd5Hash(MD5 md5Hash, string input)
        {

            // Convert the input string to a byte array and compute the hash.
            byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input));

            // Create a new Stringbuilder to collect the bytes
            // and create a string.
            StringBuilder sBuilder = new StringBuilder();

            // Loop through each byte of the hashed data 
            // and format each one as a hexadecimal string.
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }

            // Return the hexadecimal string.
            return sBuilder.ToString();
        }

        // Verify a hash against a string.
        public static bool VerifyMd5Hash(MD5 md5Hash, string input, string hash)
        {
            // Hash the input.
            string hashOfInput = GetMd5Hash(md5Hash, input);

            // Create a StringComparer an compare the hashes.
            StringComparer comparer = StringComparer.OrdinalIgnoreCase;

            if (0 == comparer.Compare(hashOfInput, hash))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}

