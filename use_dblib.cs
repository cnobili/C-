/*
 * use_dblib.cs
 *
 * Craig Nobili
 */
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using dblib;

public class UseDbLib
{
  // Private members

  private static string sqlStmt;

  public static void Main(String [] args)
  {
    sqlStmt = ""; // to avoid compiler warning about variable never used as it is commented out in places
    Console.WriteLine(sqlStmt);

    //if (args.Length != 4)
    //{
    //  Console.WriteLine("\nUsage: use_ulsd <Access DB> <Oracle DB> <Oracle Id> <Oracle Password>");
    //  return;
    //}

    //String csvFile = "c:\\plsql\\JCADW\\data\\Chartis_r2_IO_RoundsFMT_DCs_20060101-20101130_on_201012211457.txt";
    //String csvFile = "IO.txt";
    //String csvFile = "PHARMACY.txt";
    //String excelFile = "c:\\installs\\2017 YTD(OCT) - Clincial Outcomes.xlsx";
    String excelFile = "c:\\C#\\dblib\\oat.xlsx";
    //String worksheet = "Custom Query - Facility";
    String worksheet = "Epic SM Entities";
    //String cells = "A8:J27974";
    String cells = "A0:O937";
    //String tableName = "IF_CHF_CC_PHARMACY_DCS";
    //String outputFile = "if_chf_cc_pharmacy_dcs.sql";
    //String outputFile = "lab_obr16.dat";
    String csvFile = "outcomes.txt";
    String delim = "|";

    // use schema.ini file to specify non-commad delimiter, i.e. for pipe (|) delimited fields
    //dblib.DbUtil.OraGenCrTabFromCsv(csvFile, tableName, false, outputFile);
    //dblib.DbUtil.ExcelGenCsvFile2(outputFile, "|", excelFile, worksheet);
    dblib.DbUtil.ExcelGenCsvFile(csvFile, delim, excelFile, worksheet, cells);
    //dblib.DbUtil.OraGenExtTabFromExcel(excelFile, worksheet, "lab_obr16_ext.sql", "lab_obr16_ext", "|", "JMH_DW_DIR", outputFile);


  } // Main()

  public static void ReadDir(String dirPath)
  {
    DirectoryInfo dir = new DirectoryInfo(dirPath);
    FileInfo[] filesInDir = dir.GetFiles();

    foreach (FileInfo file in filesInDir)
    {
      Console.WriteLine(file.Name);
    }

  } // ReadDir()

  public static void LoadExcelFromDir(String dirPath, String tableName)
  {
    DirectoryInfo dir = new DirectoryInfo(dirPath);
    FileInfo[] filesInDir = dir.GetFiles();

    foreach (FileInfo file in filesInDir)
    {
      Console.WriteLine("Procesing file = <{0}>", file.Name);
      DbUtil.AccLoadTableFromExcel(file.Name, "ACCESS LOAD", true, tableName);
    }

  } // LoadExcelFromDir()

  public static void LoadExcelFromDir2(String dirPath, String tableName)
  {
    DirectoryInfo dir = new DirectoryInfo(dirPath);
    FileInfo[] filesInDir = dir.GetFiles();

    foreach (FileInfo file in filesInDir)
    {
      if (file.Name.StartsWith("Access Ready"))
      {
        Console.WriteLine("Procesing file = <{0}>", file.Name);
        DbUtil.AccLoadTableFromExcel4(tableName, true, dirPath + file.Name, "Access Ready");
      }
    }

  } // LoadExcelFromDir2()


  public static void LoadExcelFromExcelFile(String excelFilePath)
  {
    String excelConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" +
      "Data Source=" + excelFilePath + ";Extended Properties=" +
      "\"Excel 8.0;HDR=YES;IMEX=1;\"";

    OleDbConnection con = new OleDbConnection(excelConnect);

    con.Open();
    Console.WriteLine("Made the connection to the spreadsheet {0}", excelFilePath);

    OleDbCommand command = con.CreateCommand();
    command.CommandText = "select * from [Sheet1$]";

    OleDbDataReader reader = command.ExecuteReader();
    while (reader.Read())
    {
      Console.WriteLine("{0}:{1}:{2}:{3}", reader[0], reader[1], reader[2], reader[3]);
      DbUtil.AccLoadTableFromExcel(reader[0].ToString() + reader[1].ToString(), reader[2].ToString(), true, reader[3].ToString());
    }
      reader.Close();
      con.Close();

  } // LoadExcelFromFile

} // UseDbLib
