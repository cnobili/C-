/*
 * Csv2ExtTabOra.cs
 *
 * Generates DDL for an Oracle External Table definition.
 * For the CSV file passed in, assumes that the first record (header)
 * in the CSV file contains the column names separated by the
 * delimiter.
 *
 * Craig Nobili
 */

using System;
using System.Text.RegularExpressions;
using System.IO;

public class Csv2ExtTabOra
{

  private static String header;
  private static String filePath;
  private static String filename;
  private static String col_delimiter;
  private static String rec_delimiter;
  private static String table;
  private static String dirObj;
  private static String addRecNum;
  private static String outputDir;

  public static String padRight(String s, int n)
  {
    return String.Format("{0," + -n + "}", s);

  } // PadRight()

  public static void genTableDDL()
  {
    String colDef;
    FileInfo fi = new FileInfo(filePath);
    StreamReader sr = fi.OpenText();

    // The end of file character can be something other than carriage return/line feeed
    // We just need to get the first record which is the column header
    // Assume header record is less than 16K bytes, read this many or less bytes
    // in order to split out the header record to get the column names
    char [] buf = new char[16384];
    sr.Read(buf, 0, 16384);
    sr.Close();

    header = new String(buf);
    String [] tok = Regex.Split(header, rec_delimiter);
    header = tok[0];
    String [] tokens = header.Split(col_delimiter.ToCharArray());

    StreamWriter pw = new StreamWriter(outputDir + "\\" + table + ".sql");
    Console.WriteLine("Writing out file = <{0}>", outputDir + "\\" + table + ".sql");

    pw.WriteLine("create table " + table);
    pw.WriteLine("(");

    for (int i = 0; i < tokens.Length; i++)
    {
			if (tokens[i].Length == 0) continue;

      colDef = tokens[i].ToLower();
      colDef = Regex.Replace(colDef, @"[- ,']", "_");

      if (colDef.Length > 80)
      {
        colDef = colDef.Substring(0, 80);
      }

      colDef = padRight(colDef, 80) + " varchar(4000)";

      if (i == 0)
      {
        pw.WriteLine("  " + colDef);
      }
      else
      {
        pw.WriteLine(", " + colDef);
      }
    }

    if (addRecNum.Equals("Y"))
    {
      colDef = padRight("rec_num", 80) + " number";
      pw.WriteLine(", " + colDef);
    }

    pw.WriteLine(")");
    pw.WriteLine("organization external");
    pw.WriteLine("(");
    pw.WriteLine("  type oracle_loader");
    pw.WriteLine("  default directory " + dirObj.ToUpper());
    pw.WriteLine("  access parameters");
    pw.WriteLine("  (");
    pw.WriteLine("    records delimited by '" + rec_delimiter + "'");
    pw.WriteLine("    readsize 5000000");
    pw.WriteLine("    skip 1");
    pw.WriteLine("    nobadfile");
    pw.WriteLine("    nodiscardfile");
    pw.WriteLine("    nologfile");
    pw.WriteLine("    fields terminated by '" + col_delimiter + "'");
    pw.WriteLine("    missing field values are null");
    pw.WriteLine("    reject rows with all null fields");

    if (addRecNum .Equals("Y"))
    {
      pw.WriteLine("    (");

      for (int i = 0; i < tokens.Length; i++)
      {
        if (tokens[i].Length == 0) continue;

        colDef = tokens[i].ToLower();
        Regex.Replace(colDef, @"[- ,']", "_");

        if (colDef.Length > 30)
        {
          colDef = colDef.Substring(0, 30);
        }

        colDef = padRight(colDef, 30);

        if (i == 0)
        {
          pw.WriteLine("      " + colDef);
        }
        else
        {
          pw.WriteLine("    , " + colDef);
        }
      }

      colDef = padRight("rec_num", 30) + " recnum";
      pw.WriteLine("    , " + colDef);
      pw.WriteLine("    )");
    }

    pw.WriteLine("  )");
    pw.WriteLine("  location ('" + filename + "')");
    pw.WriteLine(")");
    pw.WriteLine("reject limit unlimited");
    pw.WriteLine("parallel");
    pw.WriteLine(";");

    pw.Close();

  } // genTableDDL()

  public static void Main(String [] args)
  {

    if (args.Length != 6)
    {
      Console.WriteLine("Usage: Csv2ExtTabOra <csv file> <col_delimiter> <rec_delimiter> <table_name> <directory object> <add rec_num Y|N>");
      return;
    }
    filePath  = args[0];
    col_delimiter = args[1];
    rec_delimiter = args[2];
    table     = args[3];
    dirObj = args[4];
    addRecNum = args[5].ToUpper();
    outputDir = Path.GetDirectoryName(filePath);
    filename = Path.GetFileName(filePath);

    genTableDDL();

  } // Main()

} // class Csv2ExtTabOra
