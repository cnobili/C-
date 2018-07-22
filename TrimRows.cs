/*
 * TrimRows.cs
 *
 * Trims blank space off the end of a record and removes blank records from file.
 *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.IO;
using System.Text;
using System.Collections;
using System.Linq;
using Microsoft.Win32;

public class TrimRows
{

  private static void GetVersionFromEnvironment()
  {
      Console.WriteLine("Version: " + Environment.Version.ToString());

  } // GetVersionFromEnvironment()

  public static void Usage()
  {
    Console.WriteLine();
    Console.WriteLine("Usage: TrimRows filename  newFilename");
    Console.WriteLine();
    Console.WriteLine("Usage: TrimRows directory");
    Console.WriteLine();
      
  } // Usage()
  
  public static String RemoveControlChars(String str)
  {
    StringBuilder sb = new StringBuilder();
  
    foreach (char c in str)
    {
      if ( (int)c >= 32 )
      {
         sb.Append(c);
      }
    }
     
    return( sb.ToString() );
  
  } // RemoveControlChars()
  
  public static void ProcessFilesInDir(String dir)
  {
  
    String[] filePaths = Directory.GetFiles(dir); 
      
    if (filePaths.Length == 0)
    {
      Console.WriteLine("No files in '{0}'", dir);
      return;
    }
    
    foreach (String fp in filePaths)
    {
      Console.WriteLine("Processing file = {0}", fp);
      TrimRecs(fp, fp + ".2");
    }
  
  } // ProcessFilesInDir()
  
  public static void TrimRecs(String filename, String newFilename)
  {
   
    using (StreamReader reader = new StreamReader(filename))
    using (StreamWriter writer = new StreamWriter(newFilename))
    {
      String line = null;
      String tmpLine = null;
      while ((line = reader.ReadLine()) != null)
      {
        tmpLine = RemoveControlChars(line);
        if (tmpLine.Equals(String.Empty)) continue;
        tmpLine = line.Replace("\r\n", "").Replace("\n", "").Replace("\r", "").Trim();
        if (tmpLine.Equals(String.Empty)) continue;
                
        writer.WriteLine(line.Trim());
      }
    }
           
  } // TrimRecs()

  public static void Main(String [] args)
  {
    GetVersionFromEnvironment();
    
    if (args.Length < 1)
    {
      Usage();
      return;
    }
        
    String filename = null;
    String newFilename = null;
    String dir = null;
    
    if (args.Length == 1)
    {
      dir = args[0];
      
      if ( !Directory.Exists(dir) )
      {
        Console.WriteLine("Error: input directory = {0} does not exist", dir);
        return;
      }
      
      ProcessFilesInDir(dir);
  
    }
    else
    {
      filename    = args[0];
      newFilename = args[1];
      
      //var lines = File.ReadAllLines(filename).Where(arg => !string.IsNullOrWhiteSpace(arg));
      //File.WriteAllLines(newFilename, lines);
          
      if ( !System.IO.File.Exists(filename) )
      {
        Console.WriteLine("filename = {0} does not exist", filename);
        return;
      }
          
      TrimRecs(filename, newFilename);
          
      //System.IO.File.Delete(filename);
      //System.IO.File.Move(newFilename, filename);

    }
    
  } // Main()

} // class TrimRows
