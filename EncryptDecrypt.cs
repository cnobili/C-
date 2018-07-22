/*
 * EncryptDecrypt.cs
 *
 * Encrypts or Decrypts a string passed in on command line.
 *
 * Craig Nobili
 */

using System;
using JMHCrypto;

public class EncryptDecrypt
{
  public static void Main(String [] args)
  {
    if (args.Length != 2)
    {
      Console.WriteLine("Usage: EncryptDecrypt <action> <string>");
      Console.WriteLine("");
      Console.WriteLine("  <action> - Encrypt|Decrypt");
      Console.WriteLine("  <string> - String to Encrypt or Decrypt");
      return;
    }
    String action = args[0];
    String str = args[1];
    
    if (action.Equals("Encrypt"))
    {
      Console.WriteLine(Crypto.Encrypt(str));
    }
    else if (action.Equals("Decrypt"))
    {
      Console.WriteLine(Crypto.Decrypt(str));
    }
    else
    {
      Console.WriteLine("Unknown action, must be either Encrypt or Decrypt");
    }
    
  } // Main()

} // EncryptDecrypt class
