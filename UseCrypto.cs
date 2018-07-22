/*
 * UseCrypto.cs
 *
 * For testing the Crypto Library.
 *
 * Craig Nobili
 */

using System;
using JMHCrypto;

public class UseCrypto
{
  public static void Main(String [] args)
  {
    Console.Write("Enter password:");
    String pwd = Crypto.ReadPassword();
    
    Console.WriteLine("\nYou entered <{0}>", pwd);
    
  } // Main()

} // UseCrypto class
