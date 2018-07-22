/*
 * Excel2Csv.cs
 *
 * Converts Excel file into a delimited flat file.
 *
 * Craig Nobili
 */
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using dblib;

public class Excel2Csv
{
  // Private members

  public static void Main(String [] args)
  {
    if (args.Length != 4)
    {
      Console.WriteLine("\nUsage: Excel2Csv <spreadsheet> <worksheet> <csv file> <delimiter>");
      return;
    }

    String excelFile = args[0];
    String worksheet = args[1];
    String csvFile   = args[2];
    String delimiter = args[3];

    // use schema.ini file to specify non-comma delimiter, i.e. for pipe (|) delimited fields

    dblib.DbUtil.ExcelGenCsvFile2(csvFile, delimiter, excelFile, worksheet);

  } // Main()

} // Excel2Csv
