using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json.Linq;
using Cryptor;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            JObject jObject = new JObject();
            jObject["HWID"] = "BFEBFBFF000306C3";
            jObject["EndDate"] = DateTime.Parse("2020-04-11");

            Console.WriteLine(Base64Encode(AES.encryptAES256(jObject.ToString(), Base64Encode("4a6d53436f6d50416e5950724f645563744b655974457354"))));
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
    }
}
