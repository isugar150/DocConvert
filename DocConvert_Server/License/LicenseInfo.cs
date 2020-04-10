using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Server.License
{
    class LicenseInfo
    {
        /// <summary>
        /// 하드웨어 아이디를 반환
        /// </summary>
        /// <returns>하드웨어 아이디</returns>
        public string getHWID()
        {
            var mbs = new ManagementObjectSearcher("Select ProcessorId From Win32_processor");
            ManagementObjectCollection mbsList = mbs.Get();
            string id = "";
            foreach (ManagementObject mo in mbsList)
            {
                id = mo["ProcessorId"].ToString();
            }
            return id;
        }

        /// <summary>
        /// AES256 복호화
        /// </summary>
        /// <param name="encryptText">암호화된 문자열</param>
        /// <param name="key">복호화할 키 값</param>
        /// <returns>Unicode 인코딩으로 복호화한 문자열</returns>
        public static string decryptAES256(string encryptText, string key)
        {
            return decryptAES256(encryptText, key, Encoding.Unicode);
        }

        /// <summary>
        /// AES256 복호화
        /// </summary>
        /// <param name="encryptText">암호화된 문자열</param>
        /// <param name="key">복호화할 키 값</param>
        /// <param name="encoding">System.Text.Encoding</param>
        /// <returns>지정된 인코딩으로 복호화한 문자열</returns>
        public static string decryptAES256(string encryptText, string key, Encoding encoding)
        {
            string Key = Base64Decoding(key);
            MemoryStream ms = null;
            CryptoStream cs = null;
            try
            {
                RijndaelManaged aes = new RijndaelManaged();

                byte[] encryptData = Convert.FromBase64String(encryptText);
                byte[] salt = Encoding.ASCII.GetBytes(Key.Length.ToString());
                PasswordDeriveBytes secretKey = new PasswordDeriveBytes(Key, salt);

                ICryptoTransform decryptor = aes.CreateDecryptor(secretKey.GetBytes(32), secretKey.GetBytes(16));
                ms = new MemoryStream(encryptData);
                cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read);
                byte[] result = new byte[encryptData.Length];
                int decryptedCount = cs.Read(result, 0, result.Length);

                return encoding.GetString(result, 0, decryptedCount);
            }
            catch (Exception e)
            {
                return "Decrypt ERROR : " + e.Message;
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

        public static string Base64Decoding(string DecodingText, System.Text.Encoding oEncoding = null)
        {
            if (oEncoding == null)
                oEncoding = System.Text.Encoding.UTF8;

            byte[] arr = System.Convert.FromBase64String(DecodingText);
            return oEncoding.GetString(arr);
        }
    }
}
