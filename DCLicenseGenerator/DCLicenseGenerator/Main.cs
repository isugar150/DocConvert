using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace DCLicenseGenerator
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Now;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string key = Base64Encode("JmDoCOnVerTerServErJmCoRp");
            string key = "JmDoCOnVerTerServErJmCoRp";
            JObject jObject = new JObject();
            jObject["HWID"] = textBox1.Text;
            jObject["EndDate"] = dateTimePicker1.Value.ToString("yyyy-MM-dd 23:59:59");
            jObject["CompanyName"] = textBox3.Text;
            Debug.WriteLine("KEY: " + key);
            Debug.WriteLine("Data: " + jObject.ToString());
            textBox2.Text = encryptAES256(jObject.ToString(), key);
        }

        /// <summary>
        /// Base64 디코딩
        /// </summary>
        /// <param name="data">디코딩할 텍스트</param>
        /// <returns></returns>
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

        /// <summary>
        /// AES256 Encrypt
        /// </summary>
        /// <param name="text">암호화할 텍스트</param>
        /// <param name="key">해제할 키</param>
        /// <returns></returns>
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
