/*
 * Crypto.cs
 *
 * Encryption/Decryption routines.
 *
 * Craig Nobili
 */

using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace JMHCrypto
{

public class Crypto
{

static readonly string PasswordHash = "P@@Sw0rd";
static readonly string SaltKey = "S@LT&KEY";
static readonly string VIKey = "@1B2c3D4e5F6g7H8";

public static string Encrypt(string plainText)
{
  byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

  byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
  var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
  var encryptor = symmetricKey.CreateEncryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));
      
  byte[] cipherTextBytes;

  using (var memoryStream = new MemoryStream())
  {
    using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
    {
      cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
      cryptoStream.FlushFinalBlock();
      cipherTextBytes = memoryStream.ToArray();
      cryptoStream.Close();
    }
    memoryStream.Close();
  }
  return Convert.ToBase64String(cipherTextBytes);
  
} // Encrypt()

public static string Decrypt(string encryptedText)
{
  byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
  byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
  var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

  var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));
  var memoryStream = new MemoryStream(cipherTextBytes);
  var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
  byte[] plainTextBytes = new byte[cipherTextBytes.Length];

  int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
  memoryStream.Close();
  cryptoStream.Close();
  return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
      
} // Decrypt()

// Mask password as it is read in from command line prompt
public static string ReadPassword()
{
  Stack<string> pass = new Stack<string>();
 
  for (ConsoleKeyInfo consKeyInfo = Console.ReadKey(true); consKeyInfo.Key != ConsoleKey.Enter; consKeyInfo = Console.ReadKey(true))
  {
    if (consKeyInfo.Key == ConsoleKey.Backspace)
    {
      try
      {
        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
        Console.Write(" ");
        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
        pass.Pop();
      }
      catch (InvalidOperationException ex)
      {
        /* Nothing to delete, go back to previous position */
        Console.SetCursorPosition(Console.CursorLeft + 1, Console.CursorTop);
      }
    }
    else
    {
      Console.Write("*");
      pass.Push(consKeyInfo.KeyChar.ToString());
    }
  }
  
  String[] password = pass.ToArray();
  Array.Reverse(password);
  return string.Join(string.Empty, password);
  
} // ReadPassword()

} // Crypto class

} // JMHCrypto namespace