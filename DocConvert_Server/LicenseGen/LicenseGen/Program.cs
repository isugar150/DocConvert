using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Diagnostics;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            //string key = Base64Encode("JmDoCOnVerTerServErJmCoRp");
            string key = "JmDoCOnVerTerServErJmCoRp";
            JObject jObject = new JObject();
            jObject["HWID"] = "BFEBFBFF000706A1";
            jObject["EndDate"] = DateTime.Parse("2021-04-11");
            Console.WriteLine("KEY: " + key);
            Console.WriteLine("Data: " + jObject.ToString());
            Console.WriteLine(encryptAES256(jObject.ToString(), key));
        }
        public static string Base64Encode(string data)
        {
            try
            {
                byte[] encData_byte = new byte[data.Length];
                encData_byte = System.Text.Encoding.UTF8.GetBytes(data);
                string encodedData = Convert.ToBase64String(encData_byte);
                return encodedData;
            }
            catch (Exception e)
            {
                throw new Exception("Error in Base64Encode: " + e.Message);
            }
        }
        public static string encryptAES256(string text, string key)
        {
            MemoryStream ms = null;
            CryptoStream cs = null;
            try
            {
                RijndaelManaged aes = new RijndaelManaged();

                byte[] textData = Encoding.Unicode.GetBytes(text);
                byte[] salt = Encoding.ASCII.GetBytes(key.Length.ToString());
                PasswordDeriveBytes secretKey = new PasswordDeriveBytes(key, salt);

                ICryptoTransform encryptor = aes.CreateEncryptor(secretKey.GetBytes(32), secretKey.GetBytes(16));
                ms = new MemoryStream();
                cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write);

                cs.Write(textData, 0, textData.Length);
                cs.FlushFinalBlock();
                return Convert.ToBase64String(ms.ToArray());
            }
            catch (Exception e)
            {
                return "Encrypt ERROR : " + e.Message;
            }
            finally
            {
                if (cs != null)
                {
                    cs.Close();
                }
                if (ms != null)
                {
                    ms.Close();
                }
            }
        }
    }
}
