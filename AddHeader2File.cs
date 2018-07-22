/*
 * AddHeader2File.cs
 *
 * Adds a header record to a flat file, columns are: column01, column02, column03 ...
 *
 * Craig Nobili
 */

using System;
using System.Data;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;

public class AddHeader2File
{

  /*
   * Private Data
   */
   
  String filename;
  String delimiter;
  String newFilename;
 
  /*
   * Public Methods
   */

  /*
   * Constructor.
   *
   * Gets database connection information from config file.
   */
  public AddHeader2File(String filename, String delimiter, String newFilename)
  {
    this.filename = filename;
    this.delimiter = delimiter;
    this.newFilename = newFilename;

  } // AddHeader2File()

  public static void Msg(String s)
	{
    Console.WriteLine("{0} => {1}", DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss"), s);

  } // Msg()

  public void AddHeader()
  {
    String[] lines = File.ReadAllLines(filename);
    String[] tokens = lines[0].Split(delimiter.ToCharArray());
    String header = null;
    String delim = null;
    
    for (int i = 0; i < tokens.Length; i++)
    {
      header = header + delim + "column" + (i + 1).ToString("D2");
      delim = delimiter;
    }
    
    // File.WriteAllLines(path, createText, Encoding.UTF8);
    
    // Need header first, use LINQ
    List<string> list = new List<string>(lines);
    StreamWriter sw = new StreamWriter(newFilename);
    sw.WriteLine(header);
    list.ForEach(r=> sw.WriteLine(r));
    sw.Close();
    
  } // AddHeader()

  /*
   * Main() - Entry point of program.
   */
  public static int Main (String [] args)
  {
    if (args.Length != 3)
    {
      Console.WriteLine("Usage:AddHeader2File <filename> <delimiter> <newFilename>");
      return(-1);
    }
    
    AddHeader2File hf = new AddHeader2File(args[0], args[1], args[2]);
        
    if (!File.Exists(hf.filename))
    {
      Msg("Error: file = <" + hf.filename + "> does not exist.");
      return(-1);
    }
    
    hf.AddHeader();
    
    return(0);

  } // Main()

} // class AddHeader2File
