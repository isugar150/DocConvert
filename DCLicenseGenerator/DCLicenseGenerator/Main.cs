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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string key = Base64Encode("JmDoCOnVerTerServErJmCoRp");
            string key = "JmDoCOnVerTerServErJmCoRp";
            JObject jObject = new JObject();
            jObject["HWID"] = textBox1.Text;
            jObject["EndDate"] = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            Debug.WriteLine("KEY: " + key);
            Debug.WriteLine("Data: " + jObject.ToString());
            new MessageDialog("라이선스 발급", "다음 코드를 DocConvert_Server.exe.config 파일에 LicenseKEY에 입력하세요.", encryptAES256(jObject.ToString(), key)).ShowDialog(this);
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
