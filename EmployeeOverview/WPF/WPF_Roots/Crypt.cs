using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace WPF_Roots
{
    public static class Crypt
    {
        private const string Röschti = "h/23M,Na9(`8]S A";
        private const string Bradwurscht = "9/hZU' Ta+78AtV.v0O0,?&a1You&6zS";

        public static string EncryptString(string ClearText, string iv = "", string key = "")
        {
            iv = (iv + Röschti).Substring(0, 16);
            key = (key + Bradwurscht).Substring(0, 32);
            byte[] clearTextBytes = Encoding.UTF8.GetBytes(ClearText);

            System.Security.Cryptography.SymmetricAlgorithm rijn =  SymmetricAlgorithm.Create();

            MemoryStream ms = new MemoryStream();
            byte[] ba_rgbIV = Encoding.ASCII.GetBytes(iv);
            byte[] ba_key = Encoding.ASCII.GetBytes(key);
            CryptoStream cs = new CryptoStream(ms, rijn.CreateEncryptor(ba_key, ba_rgbIV), CryptoStreamMode.Write);

            cs.Write(clearTextBytes, 0, clearTextBytes.Length);
            cs.Close();
            rijn.Dispose();
            rijn = null;

            return Convert.ToBase64String(ms.ToArray());
        }

        public static string DecryptString(string EncryptedText, string iv = "", string key = "")
        {
            try
            {
                iv = (iv + Röschti).Substring(0, 16);
                key = (key + Bradwurscht).Substring(0, 32);
                byte[] encryptedTextBytes = Convert.FromBase64String(EncryptedText);

                System.Security.Cryptography.SymmetricAlgorithm rijn = SymmetricAlgorithm.Create();

                MemoryStream ms = new MemoryStream();
                byte[] ba_rgbIV = Encoding.ASCII.GetBytes(iv);
                byte[] ba_key = Encoding.ASCII.GetBytes(key);
                CryptoStream cs = new CryptoStream(ms, rijn.CreateDecryptor(ba_key, ba_rgbIV), CryptoStreamMode.Write);

                cs.Write(encryptedTextBytes, 0, encryptedTextBytes.Length);
                cs.Close();
                rijn.Dispose();
                rijn = null;

                return Encoding.UTF8.GetString(ms.ToArray());
            }
            catch (Exception)
            {
                return null;
            }
        }


        /// using (MD5 md5Hash = MD5.Create())
        ///    { string hash = GetMd5Hash(md5Hash, source); }
        ///    
        public static string GetMd5Hash(MD5 md5Hash, string input)
        {
            if (input == null)
                return null;

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
        //
        public static string GetMd5Hash(string input)
        {
            using (MD5 md5Hash = MD5.Create())
                return GetMd5Hash(md5Hash, input);
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
