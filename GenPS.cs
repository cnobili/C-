/*
 * GenPS.cs
 *
 * Generate PowerShell script to dump Prime and QIP Audit/Support tables to Excel.
 *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Net;

namespace GenPS_CombineExcel
{

class GenPS
{

  /*
   * Private Data
   */
   
  public const String PROGRAM_NAME = "GenPS";
  public const String PS_SCRIPT = "zMergeExcelFiles.ps1";
  
  private static String metadataDbConnStr;
  private static String queryFile;
     
  private static OdbcConnection dbConn;
  private static String dbSqlStmt;
  
  /*
   * Public Methods
   */

  /*
   * Constructor.
   *
   */
  public GenPS(String dbSqlFile)
  {
    // Get SQL Statement
    FileInfo sqlFile = new FileInfo(dbSqlFile);
    StreamReader sr = sqlFile.OpenText();
    String line;
    dbSqlStmt = "";
    
    while ( (line = sr.ReadLine()) != null)
    {
      dbSqlStmt += line;
      dbSqlStmt += "\n";
    }
    sr.Close();
    
  } // GenPS()

  /*
   * ConnectToDb()
   *
   * Opens up a connection to the database.
   */
  public void ConnectToDb(String connStr)
  {
    Console.WriteLine("\nMake connection to database URL = " + connStr);

    // Connect to database
    dbConn = new OdbcConnection(connStr);

    try
    {
      dbConn.Open();
    }
    catch (Exception ex)
    {
      Console.WriteLine("Failed to Connect to database:" + ex.Message);
      return;
    }

  } // ConnectToDb()

  /*
   *  DisconnectFromDb()
   *
   * Disconnects from the database.
   */
  public void DisconnectFromDb(OdbcConnection dbConn)
  {
    Console.WriteLine("Disconnect from database");
    dbConn.Close();
    dbConn.Dispose();

  } // DisconnectFromDb()

  /*
   * GenFile()
   */
  public void GenFile(OdbcConnection dbConn, String dbSqlStmt)
  {
    StreamWriter pw = null;
    OdbcCommand cmdSQL = new OdbcCommand(dbSqlStmt, dbConn);
    cmdSQL.CommandTimeout = 0;
    OdbcDataReader dataReader = cmdSQL.ExecuteReader();
    int fieldCount = dataReader.FieldCount;
    
    String MetricName;
   
    Console.WriteLine("Execute SQL = " + dbSqlStmt);
    Console.WriteLine("Number of columns in select stmt = " + fieldCount);

    try
    {
      pw = new StreamWriter(PS_SCRIPT);
      
      pw.WriteLine("param");
      pw.WriteLine("(");
      pw.WriteLine("  [Parameter(Mandatory=$true)][string]$srcDir");
      pw.WriteLine(", [Parameter(Mandatory=$true)][string]$yyyymm");
      pw.WriteLine(")");
      pw.WriteLine("");
      pw.WriteLine("# ******** PROGRAMATICALLY GENERATED PowerShell Script (created by C# program GenPS) ********");
      pw.WriteLine("");
      pw.WriteLine("#Write-Host \"command line parms $srcDir and $yyyymm\"");
      pw.WriteLine("");
        
      // Iterate through the resultset
      while (dataReader.Read())
      {
        pw.WriteLine("$ExcelObject=New-Object -ComObject excel.application");
        pw.WriteLine("$ExcelObject.visible=$true");
      
        MetricName = dataReader[0].ToString();
           
        pw.WriteLine("");
        pw.WriteLine("#" + MetricName);
        pw.WriteLine("$destFilename=\"Audit_" + MetricName.Replace(' ', '_').Replace('.', '_') + "_\" + $yyyymm + \".xlsx\"");
        pw.WriteLine("$pattern=\"*" + MetricName.Replace(' ', '_').Replace('.', '_') + "*.xlsx\"");
        pw.WriteLine("");
        pw.WriteLine("$ExcelFiles=Get-ChildItem $pattern -Path $srcDir");
        pw.WriteLine("");
        pw.WriteLine("$Workbook=$ExcelObject.Workbooks.add()");
        pw.WriteLine("$Worksheet=$Workbook.Sheets.Item(\"Sheet1\")");
        pw.WriteLine("");
        pw.WriteLine("foreach($ExcelFile in $ExcelFiles)");
        pw.WriteLine("{");
        pw.WriteLine("  $Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)");
        pw.WriteLine("  $Everysheet=$Everyexcel.sheets.item(1)");
        pw.WriteLine("  $Everysheet.Copy($Worksheet)");
        pw.WriteLine("  $Everyexcel.Close()");
        pw.WriteLine("}");
        pw.WriteLine("$Workbook.SaveAs($srcDir + \"\\\" + $destFilename)");
        
        pw.WriteLine("");
        pw.WriteLine("$ExcelObject.Quit()");
      }
    }
    finally
    {
      pw.Close();
    }
    
  } // GenFile()
  
  public static void Usage()
  {
    Console.WriteLine();
    Console.WriteLine("------------------------------------------------------------------------------------------------");
    Console.WriteLine("Usage: {0} metadataDbConnStr queryFile", PROGRAM_NAME);
    Console.WriteLine();
    Console.WriteLine("  metaDbConnStr = database connection string, exampels:");
    Console.WriteLine("    --> \"Driver={SQL Server};Server=theServer;Database=theDatabase;Trusted_Connection=yes;\"");
    Console.WriteLine("    --> \"Driver={SQL Server};Server=theServer;Database=theDatabase;uid=username;pwd=password;\"");
    Console.WriteLine();
    Console.WriteLine("  queryFile = the query file to run against the metadata DB view");
    Console.WriteLine();
    Console.WriteLine("------------------------------------------------------------------------------------------------");
    Console.WriteLine();    
    
  } // Usage()
  
  /*
   * Main() - Entry point of program.
   */
  public static void Main (String [] args)
  {
    GenPS dmp = null;
   
    Console.WriteLine("Starting Program{0}, number of command line arguments = {1}\n", PROGRAM_NAME, args.Length);
    Console.WriteLine("  exe directory = {0}\n", AppDomain.CurrentDomain.BaseDirectory);
    Console.WriteLine("  current directory = {0}\n", Environment.CurrentDirectory);
    
    if (args.Length != 2)
    {
      Usage();
      return;
    }
      
    Console.WriteLine("\n{0} invoked with the following command line arguments:", PROGRAM_NAME);
    for (int i = 0; i < args.Length; i++)
    {
      Console.WriteLine("  arg {0} = {1}", i + 1, args[i]);
    }
    
    metadataDbConnStr = args[0];
    queryFile = args[1];
      
    if ( !File.Exists(queryFile) )
    {
      Console.WriteLine("Error: file = {0} not found", queryFile);
      return;
    }
      
    dmp = new GenPS(queryFile);
      
    try
    {
      dmp.ConnectToDb(metadataDbConnStr);
      dmp.GenFile(dbConn, dbSqlStmt);
    }
    catch(Exception e)
    {
      Console.WriteLine("ERROR!");
      Console.WriteLine(e.ToString());
    }
    finally
    {
      dmp.DisconnectFromDb(dbConn);
      Console.WriteLine("Done!");
    }
      
  } // Main()

} // GenPS class

} // Extract GenPS_CombineExcel
