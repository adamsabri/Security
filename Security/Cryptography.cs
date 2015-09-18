using System;
using System.Security.Cryptography;
using System.Text;

namespace Security
{
    public sealed class Cryptography
    {
        private const string Key = "sfdjf48mdfdf3054$2160amn57";

        #region MD5

        public static string Md5Encrypt(string stringToEncrypt)
        {
            if (stringToEncrypt == null)
                throw new ArgumentNullException();

            var inputBytes = Encoding.ASCII.GetBytes(stringToEncrypt);

            //generate an MD5 hash from the password. 
            //a hash is a one way encryption meaning once you generate
            //the hash, you cant derive the password back from it.
            var hashmd5 = new MD5CryptoServiceProvider();
            var pwdhash = hashmd5.ComputeHash(Encoding.ASCII.GetBytes(Key));

            // Create a new TripleDES service provider 
            var tdesProvider = new TripleDESCryptoServiceProvider {Key = pwdhash, Mode = CipherMode.ECB};

            return Convert.ToBase64String(tdesProvider.CreateEncryptor()
                                                      .TransformFinalBlock(inputBytes, 0, inputBytes.Length));
        }

        public static string Md5Decrypt(string encryptedString)
        {
            var inputBytes = Convert.FromBase64String(encryptedString);

            //generate an MD5 hash from the password. 
            //a hash is a one way encryption meaning once you generate
            //the hash, you cant derive the password back from it.
            var hashmd5 = new MD5CryptoServiceProvider();
            var pwdhash = hashmd5.ComputeHash(Encoding.ASCII.GetBytes(Key));

            // Create a new TripleDES service provider 
            var tdesProvider = new TripleDESCryptoServiceProvider {Key = pwdhash, Mode = CipherMode.ECB};

            return Encoding.ASCII.GetString(tdesProvider.CreateDecryptor()
                                                        .TransformFinalBlock(inputBytes, 0, inputBytes.Length));
        }

        #endregion
    }
}
