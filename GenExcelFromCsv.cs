/*
 * GenExcelFromCsv.cs
 *
 * OleDb Connection strings for Access and CSV files don't work for 64 bit.
 * Need 64 bit drivers,
 *
 * Or build this program as a 32 bit executable:
 *   csc GenExcelFromCsv.cs /platform:x86
 *
 * Craig Nobili
 */
using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Collections;

public class GenExcelFromCsv
{
  /*
   * Main() - Entry Point
   */
  public static void Main(String[] args)
  {
    if (args.Length != 2)
    {
      Console.WriteLine("Usage: GenExcelFromCsv csvFile excelFile\n");
      return;
    }
    
    csv2excel(args[0], args[1]);
      
  } // Main()

  /*
   * csv2excel
   *
   * Generates an Excel file from a CSV file.
   */
  public static void csv2excel(String csvFile, String excelFile)
  {
    String excelWorksheet = Path.GetFileNameWithoutExtension(csvFile);
    Console.WriteLine("Worksheet name = " + excelWorksheet);
    
    String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
      "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
 
    OleDbConnection con = new OleDbConnection(csvConnect);

    con.Open();
    Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

    OleDbCommand cmd = con.CreateCommand();
    cmd.CommandText = "SELECT * INTO " + "[" + excelWorksheet + "] IN '" + excelFile + "' 'Excel 8.0;'" + " FROM [" + csvFile + "]";
    cmd.ExecuteNonQuery();
      
    con.Close();
   
  } // csv2excel()

} // class GenExcelFromCsv
