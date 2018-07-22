/*
 * Csv2ExtTab.cs
 *
 * Generates the DDL for an Oracle external table defined against
 * the flat file.
 *
 * Craig Nobili
 */
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using dblib;

public class Csv2ExtTab
{
  // Private members

  public static void Main(String [] args)
  {
    if (args.Length != 5)
    {
      Console.WriteLine("\nUsage: Csv2ExtTab <csv file> <sql file> <oracle external table> <delimiter> <oracle directory object>");
      return;
    }

    String csvFile   = args[0];
    String sqlFile   = args[1];
    String extTable  = args[2];
    String delimiter = args[3];
    String dirObj    = args[4];

    // use schema.ini file to specify non-comma delimiter, i.e. for pipe (|) delimited fields
    if (delimiter != ",")
      genSchemaINI(csvFile, delimiter);

    dblib.DbUtil.OraGenExtTabFromCsv(csvFile, sqlFile, extTable, delimiter, dirObj);

  } // Main()

  public static void genSchemaINI(String csvFile, String delimiter)
  {
    StreamWriter p = new StreamWriter("schema.ini");

    p.WriteLine("[" + csvFile + "]");
    p.WriteLine("Format=Delimited(" + delimiter + ")");

    p.Close();

  } // genSchemaINI()

} // Csv2ExtTab
