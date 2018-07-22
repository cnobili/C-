/*
 * Excel2CsvExtTab.cs
 *
 * Converts Excel file into delimited flat file and also
 * generates the DDL for an Oracle external table defined against
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

public class Excel2CsvExtTab
{
  // Private members

  public static void Main(String [] args)
  {
    if (args.Length != 6)
    {
      Console.WriteLine("\nUsage: Excel2CsvExtTab <spreadsheet> <worksheet> <csv file> <delimiter> <oracle directory object> <oracle external table>");
      return;
    }

    String excelFile = args[0];
    String worksheet = args[1];
    String csvFile   = args[2];
    String delimiter = args[3];
    String dirObj    = args[4];
    String extTable  = args[5];

    // use schema.ini file to specify non-comma delimiter, i.e. for pipe (|) delimited fields

    dblib.DbUtil.ExcelGenCsvFile2(csvFile, delimiter, excelFile, worksheet);
    dblib.DbUtil.OraGenExtTabFromExcel(excelFile, worksheet, extTable + ".sql", extTable, delimiter, dirObj, csvFile);

  } // Main()

} // Excel2CsvExtTab
