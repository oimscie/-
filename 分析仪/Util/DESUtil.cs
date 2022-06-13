using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Gama.Util
{
    class DESUtil
    {
        #region DES对称加密解密,密匙“qwertyuiop”
        /// <summary> 加密字符串
        /// </summary> 
        /// <param name="strText">需被加密的字符串</param> 
        /// <param name="strEncrKey">密钥</param> 
        /// <returns></returns> 
        public static string DesEncrypt(string strText, string strEncrKey)
        {
            try
            {
                byte[] byKey = null;
                byte[] IV = { 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF };
                byKey = Encoding.UTF8.GetBytes(strEncrKey.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                byte[] inputByteArray = Encoding.UTF8.GetBytes(strText);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                return Convert.ToBase64String(ms.ToArray());
            }
            catch
            {
                return "";
            }
        }
        /// <summary> 解密字符串
        /// </summary> 
        /// <param name="strText">需被解密的字符串</param> 
        /// <param name="sDecrKey">密钥</param> 
        /// <returns></returns> 
        public static string DesDecrypt(string strText, string sDecrKey)
        {
            try
            {
                byte[] byKey = null;
                byte[] IV = { 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF };
                byte[] inputByteArray = new Byte[strText.Length];
                byKey = Encoding.UTF8.GetBytes(sDecrKey.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                inputByteArray = Convert.FromBase64String(strText);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                Encoding encoding = new UTF8Encoding();
                return encoding.GetString(ms.ToArray());
            }
            catch
            {
                return null;
            }
        }
        #endregion
    }
}
