/*
 * DataTools.cs
 *
 * Program that provides various data extract tools for use.
 * Uses library class DataLib.cs.
 *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;

using DataLib;

class DataTools
{
  public const String PROGRAM_NAME = "DataTools";
  
  private const String ACTION_GEN_ORA_EXT_TAB_FROM_FILE          = "GenOraExtTabFromFile";
  private const String ACTION_BCP_EXTRACT_DATA                   = "BcpExtractData";
  private const String ACTION_BCP_LOAD                           = "BcpLoad";
  private const String ACTION_GEN_DDL_BULKINSERT_STMT            = "Csv2BulkInsert";
  private const String ACTION_BULK_COPY_TABLE                    = "SqlBulkCopyTable";
  private const String ACTION_BULKINSERT_QUERY2TABLE             = "SqlBulkLoadQuery2Table";
  private const String ACTION_SQL_EXTRACT_DATA                   = "SqlExtractData";
  private const String ACTION_ODBC_EXTRACT_DATA                  = "OdbcExtractData";
  private const String ACTION_ODBC_EXTRACT_DATA_RM_CONTROL_CHARS = "OdbcExtractDataRemoveControlChars";
  private const String ACTION_ODBC_EXTRACT_DATA_REPLACE_NEWLINES = "OdbcExtractDataReplaceNewLines";
  private const String ACTION_ADD_HEADER2FILE                    = "AddHeader2File";
  private const String ACTION_GEN_TSQL_FORMAT_FILE               = "GenTsqlFormatFile";
  private const String ACTION_GEN_TSQL_READ_FILE_AS_TABLE        = "GenTsqlReadFileAsTable";
  private const String ACTION_GEN_SELECT_STMT                    = "GenSelectStmt";
  private const String ACTION_EXCEL2CSV                          = "Excel2Csv";
  private const String ACTION_EXCEL2CSV_CELLS                    = "Excel2CsvCells";
  private const String ACTION_EXCEL2CSV_RANGE                    = "Excel2CsvNamedRange";
  private const String ACTION_EXCEL2CSV_SQL                      = "Excel2CsvSql";
  private const String ACTION_ADD_AUDIT_INFO_TO_FILE             = "AddAuditInfoToFile";
  private const String ACTION_ADD_AUDIT_INFO_TO_FILES_IN_DIR     = "AddAuditInfoToFilesInDir";
  private const String ACTION_ODBC2SQL                           = "Odbc2Sql";
  private const String ACTION_CREATE_SQL_TABLE_FROM_FILE         = "CreateSqlTableFromFile";
  private const String ACTION_CSV2TABLE                          = "Csv2Table";
  private const String ACTION_FILE2TABLE                         = "File2Table";
  private const String ACTION_EXCEL2TABLE                        = "Excel2Table";
  private const String ACTION_EXCEL2TABLE_ADD_AUDIT_COLS         = "Excel2TableAddAuditCols";
  private const String ACTION_EXCEL_FILES_IN_DIR2TABLE           = "ExcelFilesInDir2Table";
  private const String ACTION_FILES_IN_DIR2TABLE                 = "FilesInDir2Table";
  private const String ACTION_GEN_EXCEL_FROM_MSSQL               = "GenExcelFromMSSQL";
  private const String ACTION_GEN_EXCEL_FROM_MSSQL_QUERY         = "GenExcelFromMSSQL_FromQuery";
  
  public static void Usage()
  {
    Console.WriteLine();
    Console.WriteLine("------------------------------------------------------------------------------------------------");
    Console.WriteLine("Usage: {0} action arg2 arg3 arg4 ... argN", PROGRAM_NAME);
    Console.WriteLine();
    Console.WriteLine("  where arg1 = action");
    Console.WriteLine("  followed by one or more other arguments depending on the action");
    Console.WriteLine();
    Console.WriteLine("EXAMPLES:");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filePath colDelimiter recDelimiter table dirObj addRecNum", PROGRAM_NAME, ACTION_GEN_ORA_EXT_TAB_FROM_FILE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} outputfile delimiter query server database user pass", PROGRAM_NAME, ACTION_BCP_EXTRACT_DATA);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} server database schema table datafile fieldDelimiter outputDir firstRow maxErrors batchSize appendOrTruncate user pass", PROGRAM_NAME, ACTION_BCP_LOAD);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} csvFile colDelimiter recDelimiter tableName", PROGRAM_NAME, ACTION_GEN_DDL_BULKINSERT_STMT);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} srcConnStr srcTable dstConnStr dstTable batchSize", PROGRAM_NAME, ACTION_BULK_COPY_TABLE);
    Console.WriteLine();
    Console.WriteLine("  Example:\n  DataTools SqlBulkCopyTable \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" sourceTable \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" destinationTable 500000");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} srcConnStr srcQueryFile dstConnStr dstTable batchSize", PROGRAM_NAME, ACTION_BULKINSERT_QUERY2TABLE);
    Console.WriteLine();
    Console.WriteLine("  Example:\n  DataTools SqlBulkLoadQuery2Table \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" srcQueryFile.sql \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" destinationTable 0");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} outputFile delimiter columnHeader(Y|N) fieldsInQuotes(T|F) queryFile server database [user] [pass]", PROGRAM_NAME, ACTION_SQL_EXTRACT_DATA);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} outputFile delimiter columnHeader(Y|N) fieldsInQuotes(T|F) queryFile connectStr", PROGRAM_NAME, ACTION_ODBC_EXTRACT_DATA);
    Console.WriteLine();
    Console.WriteLine("  Example:\n  DataTools OdbcExtractData out.txt , Y F qry.sql \"Driver={SQL Server};Server=theServer;Database=theDatabase;Trusted_Connection=yes;\"");
    Console.WriteLine("  DataTools OdbcExtractData out.txt , Y F qry.sql \"Driver={SQL Server};Server=theServer;Database=theDatabase;uid=username;pwd=password;\"");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} outputFile delimiter columnHeader(Y|N) queryFile connectStr", PROGRAM_NAME, ACTION_ODBC_EXTRACT_DATA_RM_CONTROL_CHARS);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} outputFile delimiter columnHeader(Y|N) queryFile connectStr replaceChar(s)", PROGRAM_NAME, ACTION_ODBC_EXTRACT_DATA_REPLACE_NEWLINES);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filename delimiter newFilename", PROGRAM_NAME, ACTION_ADD_HEADER2FILE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filename columnDelimiter rowDelimiter", PROGRAM_NAME, ACTION_GEN_TSQL_FORMAT_FILE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filename columnDelimiter rowDelimiter", PROGRAM_NAME, ACTION_GEN_TSQL_READ_FILE_AS_TABLE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} outputfile schemaName tableName server database [user] [pass]", PROGRAM_NAME, ACTION_GEN_SELECT_STMT);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} csvFile includeHeaderFlg(T|F) delimiter excelFile worksheet worksheetHasHeaderFlg(T|F)", PROGRAM_NAME, ACTION_EXCEL2CSV);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} csvFile includeHeaderFlg(T|F) delimiter excelFile worksheet worksheetHasHeaderFlg(T|F) cells", PROGRAM_NAME, ACTION_EXCEL2CSV_CELLS);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} csvFile includeHeaderFlg(T|F) delimiter excelFile worksheet worksheetHasHeaderFlg(T|F) namedRange", PROGRAM_NAME, ACTION_EXCEL2CSV_RANGE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} csvFile includeHeaderFlg(T|F) delimiter excelFile worksheet worksheetHasHeaderFlg(T|F) sqlStmt", PROGRAM_NAME, ACTION_EXCEL2CSV_SQL);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} srcFilePath dstFilePath fieldDelimiter firstRow", PROGRAM_NAME, ACTION_ADD_AUDIT_INFO_TO_FILE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} srcDir filePattern dstDir fieldDelimiter firstRow", PROGRAM_NAME, ACTION_ADD_AUDIT_INFO_TO_FILES_IN_DIR);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} odbcConnStr queryFile sqlConnStr schema tableName truncate|append", PROGRAM_NAME, ACTION_ODBC2SQL);
    Console.WriteLine();
    Console.WriteLine("  Example:\n  DataTools Odbc2Sql \"DRIVER={KB_SQL ODBC 32-bit Driver};UID=userId;PWD=password;TCP_PORT=5545;HOST=hostName;\" queryFile.sql \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" schema destTableName truncate");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filename delim schema table sqlConnStr", PROGRAM_NAME, ACTION_CREATE_SQL_TABLE_FROM_FILE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filename header(T|F) schema table Truncate|Append sqlConnStr", PROGRAM_NAME, ACTION_CSV2TABLE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} filename delim header(T|F) schema table Truncate|Append sqlConnStr", PROGRAM_NAME, ACTION_FILE2TABLE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} excelFile worksheet header(T|F) schema table Truncate|Append sqlConnStr", PROGRAM_NAME, ACTION_EXCEL2TABLE);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} excelFile worksheet header(T|F) schema table Truncate|Append sqlConnStr", PROGRAM_NAME, ACTION_EXCEL2TABLE_ADD_AUDIT_COLS);
    Console.WriteLine();
    Console.WriteLine(">{0} {1} directory filePattern worksheet header(T|F) schema table Truncate|Append sqlConnStr", PROGRAM_NAME, ACTION_EXCEL_FILES_IN_DIR2TABLE);
    Console.WriteLine();
    Console.WriteLine("  Note: the worksheet name, format, and header (if present) needs to be the same in all Excel files in directory");
    Console.WriteLine("  Note: the table has to exist already in the database");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} directory filePattern header(T|F) schema table Truncate|Append sqlConnStr", PROGRAM_NAME, ACTION_FILES_IN_DIR2TABLE);
    Console.WriteLine();
    Console.WriteLine("  Note: the all files in directory must be in the same format");
    Console.WriteLine("  Note: the table has to exist already in the database");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} oleDbConnectStr selectStmt excelFile", PROGRAM_NAME, ACTION_GEN_EXCEL_FROM_MSSQL);
    Console.WriteLine();
    Console.WriteLine("  Example:\n  DataTools GenExcelFromMSSQL \"Provider=SQLOLEDB;Data Source=theServer;Initial Catalog=theDB;Integrated Security=SSPI;\" \"select * from foobar \" excelFile");
    Console.WriteLine();
    Console.WriteLine(">{0} {1} oleDbConnectStr queryFile excelFile", PROGRAM_NAME, ACTION_GEN_EXCEL_FROM_MSSQL_QUERY);
    Console.WriteLine();
    Console.WriteLine("  Example:\n  DataTools GenExcelFromMSSQL_FromQuery \"Provider=SQLOLEDB;Data Source=theServer;Initial Catalog=theDB;Integrated Security=SSPI;\" query.sql excelFile");
    Console.WriteLine();
    Console.WriteLine("------------------------------------------------------------------------------------------------");
    Console.WriteLine();    
    
  } // Usage()

  /*
   * Main() - Entry point of program.
   */
  public static void Main (String [] args)
  {
    String action;
   
    Console.WriteLine("Starting Program{0}, number of command line arguments = {1}", PROGRAM_NAME, args.Length);
    
    if (args.Length == 0)
    {
      Usage();
      return;
    }
    
    Console.WriteLine("\n{0} invoked with the following command line arguments:", PROGRAM_NAME);
    action = args[0];
    for (int i = 0; i < args.Length; i++)
    {
      Console.WriteLine("  arg {0} = {1}", i + 1, args[i]);
    }
    
    //
    // Call handler routine based on action passed in
    //
    if ( action.ToUpper().Equals(ACTION_GEN_ORA_EXT_TAB_FROM_FILE.ToUpper()) )
    {
      if (args.Length != 7)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filePath      = args[1];;
      String filename      = Path.GetFileName(filePath);
      String colDelimiter  = args[2];
      String recDelimiter  = args[3];
      String table         = args[4];
      String dirObj        = args[5];;
      String addRecNum     = args[6].ToUpper();;
      String outputDir     = Path.GetDirectoryName(filePath);
      
      Console.WriteLine("Generate Oracle External Table Definition");
      DataLib.DataUtil.OraGenExtTableDDL(filePath, filename, colDelimiter, recDelimiter, table, dirObj, addRecNum, outputDir);
    }
    else if ( action.ToUpper().Equals(ACTION_BCP_EXTRACT_DATA.ToUpper()) )
    {
      if ( !(args.Length >= 6 && args.Length <= 8) )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }

      String outputfile = args[1];
      String delimiter  = args[2];
      String query      = args[3];
      String server     = args[4];
      String database   = args[5];
      String user       = null;
      String pass       = null;
                  
      if (args.Length > 6)
      {
        user = args[6];
        pass = args[7];
      }
      
      Console.WriteLine("Extract data to = {0}", outputfile);
      DataLib.DataUtil.BcpExtractData(outputfile, delimiter, query, server, database, user, pass);
    }
    else if ( action.ToUpper().Equals(ACTION_BCP_LOAD.ToUpper()) )
    {
      if ( !(args.Length >= 12 && args.Length <= 14) )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
        
      String server           = args[1];
      String database         = args[2];
      String schema           = args[3];
      String table            = args[4];
      String filePath         = args[5];
      String delimiter        = args[6];
      String outputDir        = args[7];
      int    firstRow         = Convert.ToInt32(args[8]);
      long   maxErrors        = Convert.ToInt64(args[9]);
      long   batchSize        = Convert.ToInt64(args[10]);
      String appendOrTruncate = args[11];
      String user             = null;
      String pass             = null;
      
      if (args.Length > 12)
      {
        user = args[12];
        pass = args[13];
      }
  
      Console.WriteLine("BCP Extract data");
      DataLib.DataUtil.BcpLoad(server, database, schema, table, filePath, delimiter, outputDir, firstRow, maxErrors, batchSize, user , pass, appendOrTruncate);
    }
    else if ( action.ToUpper().Equals(ACTION_GEN_DDL_BULKINSERT_STMT.ToUpper()) )
    {
      if ( args.Length != 5 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filePath      = args[1]; 
      String col_delimiter = args[2]; 
      String rec_delimiter = args[3];
      String table         = args[4];
      String outputDir     = Path.GetDirectoryName(filePath);
      
      Console.WriteLine("Bulk Insert statement creation with Table DDL");
      DataLib.DataUtil.Csv2BulkInsert(filePath, col_delimiter, rec_delimiter, table, outputDir);
    }
    else if ( action.ToUpper().Equals(ACTION_BULK_COPY_TABLE.ToUpper()) )
    {
      if ( args.Length != 6 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String srcConnStr = args[1];
      String srcTable   = args[2];
      String dstConnStr = args[3];
      String dstTable   = args[4];
      int    batchSize  = Convert.ToInt32(args[5]);
      
      Console.WriteLine("Bulk Copy Source Table = {0} to destination Table = {1}", srcTable, dstTable);
      DataLib.DataUtil.BulkCopyTable(srcConnStr, srcTable, dstConnStr, dstTable, batchSize);
    }
    else if ( action.ToUpper().Equals(ACTION_BULKINSERT_QUERY2TABLE.ToUpper()) )
    {
      if ( args.Length != 6 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String srcConnStr   = args[1];
      String srcQueryFile = args[2];
      String dstConnStr   = args[3];
      String dstTable     = args[4];
      int    batchSize    = Convert.ToInt32(args[5]);
      
      Console.WriteLine("Bulk Load destination Table = {0} from source query file = {1}", dstTable, srcQueryFile);
      String srcQuery = DataLib.DataUtil.FileToString(srcQueryFile);
      DataLib.DataUtil.BulkLoadQuery2Table(srcConnStr, srcQuery, dstConnStr, dstTable, batchSize);
    }
    else if ( action.ToUpper().Equals(ACTION_SQL_EXTRACT_DATA.ToUpper()) )
    {
      if ( !(args.Length >= 8 && args.Length <= 10) )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String outputFile     = args[1];
      String delimiter      = args[2];
      String columnHeader   = args[3];
      String fieldsInQuotes = args[4];
      String queryFile      = args[5];
      String server         = args[6];
      String database       = args[7];
      String user           = null;
      String pass           = null;
      
      Boolean fieldsInQuotesFlg = false;
      if (fieldsInQuotes.ToUpper().Equals("T"))
        fieldsInQuotesFlg = true;
      
      if (args.Length > 8)
      {
        user = args[8];
        pass = args[9];
      }
      
      Console.WriteLine("MS SQL Extract data, outputFile = {0}", outputFile);
      DataLib.DataUtil.SqlExtractData(outputFile, delimiter, columnHeader, fieldsInQuotesFlg, queryFile, server, database, user, pass);
    }
    else if ( action.ToUpper().Equals(ACTION_ODBC_EXTRACT_DATA.ToUpper()) )
    {
      if ( args.Length != 7 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String outputFile     = args[1];
      String delimiter      = args[2];
      String columnHeader   = args[3];
      String fieldsInQuotes = args[4];
      String queryFile      = args[5];
      String connStr        = args[6];
      
      Boolean fieldsInQuotesFlg = false;
      if (fieldsInQuotes.ToUpper().Equals("T"))
        fieldsInQuotesFlg = true;
      
      Console.WriteLine("ODBC Extract data, outputFile = {0}", outputFile);
      String query = DataLib.DataUtil.FileToString(queryFile);
      Boolean columnHeaderFlg = columnHeader.ToUpper() == "Y" ? true : false;
      DataLib.DataUtil.OdbcExtractData(connStr, query, delimiter, outputFile, columnHeaderFlg, fieldsInQuotesFlg);
    }
    else if ( action.ToUpper().Equals(ACTION_ODBC_EXTRACT_DATA_RM_CONTROL_CHARS.ToUpper()) )
    {
      if ( args.Length != 6 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String outputFile   = args[1];
      String delimiter    = args[2];
      String columnHeader = args[3];
      String queryFile    = args[4];
      String connStr      = args[5];
      
      Console.WriteLine("ODBC Extract data, outputFile = {0}", outputFile);
      String query = DataLib.DataUtil.FileToString(queryFile);
      Boolean columnHeaderFlg = columnHeader.ToUpper() == "Y" ? true : false;
      DataLib.DataUtil.OdbcExtractDataRemoveControlChars(connStr, query, delimiter, outputFile, columnHeaderFlg);
    }
    else if ( action.ToUpper().Equals(ACTION_ODBC_EXTRACT_DATA_REPLACE_NEWLINES.ToUpper()) )
    {
      if ( args.Length != 7 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String outputFile   = args[1];
      String delimiter    = args[2];
      String columnHeader = args[3];
      String queryFile    = args[4];
      String connStr      = args[5];
      String replaceChars = args[6];
      
      Console.WriteLine("ODBC Extract data, outputFile = {0}", outputFile);
      String query = DataLib.DataUtil.FileToString(queryFile);
      Boolean columnHeaderFlg = columnHeader.ToUpper() == "Y" ? true : false;
      DataLib.DataUtil.OdbcExtractDataReplaceNewLines(connStr, query, delimiter, outputFile, columnHeaderFlg, replaceChars);
    }
    else if ( action.ToUpper().Equals(ACTION_ADD_HEADER2FILE.ToUpper()) )
    {
      if ( args.Length != 4 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filename     = args[1];
      String delimiter    = args[2];
      String newFilename  = args[3];
      
      Console.WriteLine("Adding header record to filename = {0}, columns are column01, column02 ... columnN",newFilename);
      
      if (!File.Exists(filename))
      {
        DataLib.DataUtil.Msg("Error: file = <" + filename + "> does not exist.");
        return;
      }

      DataLib.DataUtil.AddHeader2File(filename, delimiter, newFilename);
    }
    else if ( action.ToUpper().Equals(ACTION_GEN_TSQL_FORMAT_FILE.ToUpper()) )
    {
      if ( args.Length != 4 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filename     = args[1];
      String colDelimiter = args[2];
      String rowDelimiter = args[3];
      
      Console.WriteLine("Generating BCP format file");
      
      if (!File.Exists(filename))
      {
        DataLib.DataUtil.Msg("Error: file = <" + filename + "> does not exist.");
        return;
      }

      DataLib.DataUtil.GenTsqlFormatFile(filename, colDelimiter, rowDelimiter);
    }
    else if ( action.ToUpper().Equals(ACTION_GEN_TSQL_READ_FILE_AS_TABLE.ToUpper()) )
    {
      if ( args.Length != 4 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filename     = args[1];
      String colDelimiter = args[2];
      String rowDelimiter = args[3];
      
      Console.WriteLine("Generate read file as table T-SQL statement");
      
      if (!File.Exists(filename))
      {
        DataLib.DataUtil.Msg("Error: file = <" + filename + "> does not exist.");
        return;
      }
           
      DataLib.DataUtil.GenTsqlReadFileAsTable(filename, colDelimiter, rowDelimiter);
    }
    else if ( action.ToUpper().Equals(ACTION_GEN_SELECT_STMT.ToUpper()) )
    {
      if ( !(args.Length >= 6 && args.Length <= 8) )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String outputFile   = args[1];
      String schema       = args[2];
      String table        = args[3];
      String server       = args[4];
      String database     = args[5];
      String user         = null;
      String pass         = null;
      
      if (args.Length > 6)
      {
        user = args[6];
        pass = args[7];
      }
      
      Console.WriteLine("Writing out SQL Statement to outputFile = {0}", outputFile);
      DataLib.DataUtil.GenSelectStmt(outputFile, schema, table, server, database, user , pass);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL2CSV.ToUpper()) )
    {
      if ( args.Length != 7 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String csvFile               = args[1];
      String includeHeaderFlg      = args[2];
      String delimiter             = args[3];
      String excelFile             = args[4];
      String worksheet             = args[5];
      String worksheetHasHeaderFlg = args[6];
                  
      if (!File.Exists(excelFile))
      {
        DataLib.DataUtil.Msg("Error: file = <" + excelFile + "> does not exist.");
        return;
      }
           
      DataLib.DataUtil.Excel2Csv(csvFile, includeHeaderFlg, delimiter, excelFile, worksheet, worksheetHasHeaderFlg);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL2CSV_RANGE.ToUpper()) )
    {
      if ( args.Length != 8 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String csvFile               = args[1];
      String includeHeaderFlg      = args[2];
      String delimiter             = args[3];
      String excelFile             = args[4];
      String worksheet             = args[5];
      String worksheetHasHeaderFlg = args[6];
      String namedRange            = args[7];
                  
      if (!File.Exists(excelFile))
      {
        DataLib.DataUtil.Msg("Error: file = <" + excelFile + "> does not exist.");
        return;
      }
           
      DataLib.DataUtil.Excel2CsvRange(csvFile, includeHeaderFlg, delimiter, excelFile, worksheet, worksheetHasHeaderFlg, namedRange);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL2CSV_CELLS.ToUpper()) )
    {
      if ( args.Length != 8 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String csvFile               = args[1];
      String includeHeaderFlg      = args[2];
      String delimiter             = args[3];
      String excelFile             = args[4];
      String worksheet             = args[5];
      String worksheetHasHeaderFlg = args[6];
      String cells                 = args[7];
                  
      if (!File.Exists(excelFile))
      {
        DataLib.DataUtil.Msg("Error: file = <" + excelFile + "> does not exist.");
        return;
      }
           
      DataLib.DataUtil.Excel2CsvCells(csvFile, includeHeaderFlg, delimiter, excelFile, worksheet, worksheetHasHeaderFlg, cells);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL2CSV_SQL.ToUpper()) )
    {
      if ( args.Length != 8 )
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String csvFile               = args[1];
      String includeHeaderFlg      = args[2];
      String delimiter             = args[3];
      String excelFile             = args[4];
      String worksheet             = args[5];
      String worksheetHasHeaderFlg = args[6];
      String sqlStmt               = args[7];
                  
      if (!File.Exists(excelFile))
      {
        DataLib.DataUtil.Msg("Error: file = <" + excelFile + "> does not exist.");
        return;
      }
           
      DataLib.DataUtil.Excel2CsvSql(csvFile, includeHeaderFlg, delimiter, excelFile, worksheet, worksheetHasHeaderFlg, sqlStmt);
    }
    else if ( action.ToUpper().Equals(ACTION_ADD_AUDIT_INFO_TO_FILE.ToUpper()) )
    {
      if (args.Length != 5)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filePath   = args[1];
      String outputfile = args[2];
      String delimiter  = args[3];
      int firstRow      = Convert.ToInt32(args[4]);
      
      Console.WriteLine("Input file = {0}, adding audit information to output file = {1}", filePath, outputfile);
      DataLib.DataUtil.AddAuditInfoToFile(filePath, outputfile, delimiter, firstRow);
    }
    else if ( action.ToUpper().Equals(ACTION_ADD_AUDIT_INFO_TO_FILES_IN_DIR.ToUpper()) )
    {
      if (args.Length != 6)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String inputDir    = args[1];
      String filePattern = args[2];
      String outputDir   = args[3];
      String delimiter   = args[4];
      int firstRow       = Convert.ToInt32(args[5]);
      
      Console.WriteLine("Adding audit information to files in dir = {0} matching pattern = {1} firstRow = {2}", inputDir, filePattern, firstRow);
      DataLib.DataUtil.AddAuditInfoToFilesInDir(inputDir, filePattern, outputDir, delimiter, firstRow);
    }
    else if ( action.ToUpper().Equals(ACTION_ODBC2SQL.ToUpper()) )
    {
      if (args.Length != 7)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String odbcConnStr      = args[1];
      String queryFile        = args[2];
      String sqlConnStr       = args[3];
      String schemaName       = args[4];
      String tableName        = args[5];
            
      Boolean truncateFlg = args[6].ToUpper() == "TRUNCATE" ? true : false;
      String query = DataLib.DataUtil.FileToString(queryFile);
            
      Console.WriteLine("Load SQL Server table = {0} from Odbc database", tableName);
      DataLib.DataUtil.OdbcQuery2SqlTable(odbcConnStr, query, sqlConnStr, schemaName, tableName, truncateFlg);
    }
    else if ( action.ToUpper().Equals(ACTION_CREATE_SQL_TABLE_FROM_FILE.ToUpper()) )
    {
      if (args.Length != 6)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String filename         = args[1];
      String delim            = args[2];
      String schema           = args[3];
      String tableName        = args[4];
      String sqlConnStr       = args[5];
            
      Console.WriteLine("Create SQL Table = {0} from file = {1}", tableName, filename);
      DataLib.DataUtil.CreateTable(schema, tableName, sqlConnStr, filename, delim);
    }
    else if ( action.ToUpper().Equals(ACTION_CSV2TABLE.ToUpper()) )
    {
      if (args.Length != 7)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String csvFile          = args[1];
      String header           = args[2];
      String schema           = args[3];
      String tableName        = args[4];
      String truncateOrAppend = args[5];
      String sqlConnStr       = args[6];
      
      String hdr = null;
      
      if (header.ToUpper().Equals("T"))
        hdr = "YES";
      else 
        hdr = "NO";
            
      if ( hdr.Equals("YES") && !DataLib.DataUtil.TableExists(schema, tableName, sqlConnStr) )
      {
        DataLib.DataUtil.CreateTable(schema, tableName, sqlConnStr, csvFile, ",");
      }
      if (truncateOrAppend.ToUpper().Equals("TRUNCATE"))
      {
        DataLib.DataUtil.TruncateTable(sqlConnStr, schema, tableName);
      }
            
      Console.WriteLine("Load table = {0} from file = {1}", tableName, csvFile);
      DataLib.DataUtil.File2Table(csvFile, hdr, schema, tableName, sqlConnStr);
    }
    else if ( action.ToUpper().Equals(ACTION_FILE2TABLE.ToUpper()) )
    {
      if (args.Length != 8)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String flatFile         = args[1];
      String delim            = args[2];
      String header           = args[3];
      String schema           = args[4];
      String tableName        = args[5];
      String truncateOrAppend = args[6];
      String sqlConnStr       = args[7];
      
      String hdr = null;
      
      if (header.ToUpper().Equals("T"))
        hdr = "YES";
      else 
        hdr = "NO";
      
      if ( hdr.Equals("YES") && !DataLib.DataUtil.TableExists(schema, tableName, sqlConnStr) )
      {
        DataLib.DataUtil.CreateTable(schema, tableName, sqlConnStr, flatFile, delim);
      }
      if (truncateOrAppend.ToUpper().Equals("TRUNCATE"))
      {
        DataLib.DataUtil.TruncateTable(sqlConnStr, schema, tableName);
      }
      
      Console.WriteLine("Load table = {0} from file = {1}", tableName, flatFile);
      
      String dir = Path.GetDirectoryName(flatFile);
      String filename = Path.GetFileName(flatFile);
      String fullPathSchemaINI = "schema.ini";
      
      if ( !(dir == null || dir.Equals("")) )
      {
        fullPathSchemaINI = dir + Path.DirectorySeparatorChar + "schema.ini";
      }
      DataLib.DataUtil.GenSchemaINI(fullPathSchemaINI, filename, delim, (hdr.ToUpper() == "YES" ? "True" : "False") );
      DataLib.DataUtil.File2Table(flatFile, hdr, schema, tableName, sqlConnStr);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL2TABLE.ToUpper()) )
    {
      if (args.Length != 8)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String excelFile        = args[1];
      String worksheet        = args[2];
      String header           = args[3];
      String schema           = args[4];
      String tableName        = args[5];
      String truncateOrAppend = args[6];
      String sqlConnStr       = args[7];
      
      String hdr = null;
      
      if (header.ToUpper().Equals("T"))
        hdr = "YES";
      else 
        hdr = "NO";
      
      if ( hdr.Equals("YES") && !DataLib.DataUtil.TableExists(schema, tableName, sqlConnStr) )
      {
        DataLib.DataUtil.CreateTableFromExcel(schema, tableName, sqlConnStr, excelFile, worksheet, hdr);
      }
      if (truncateOrAppend.ToUpper().Equals("TRUNCATE"))
      {
        DataLib.DataUtil.TruncateTable(sqlConnStr, schema, tableName);
      }
      
      Console.WriteLine("Load table = {0} from file = {1}", tableName, excelFile);
     
      DataLib.DataUtil.Excel2Table(excelFile, worksheet, hdr, schema, tableName, sqlConnStr);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL2TABLE_ADD_AUDIT_COLS.ToUpper()) )
    {
      if (args.Length != 8)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String excelFile        = args[1];
      String worksheet        = args[2];
      String header           = args[3];
      String schema           = args[4];
      String tableName        = args[5];
      String truncateOrAppend = args[6];
      String sqlConnStr       = args[7];
      
      String hdr = null;
      
      if (header.ToUpper().Equals("T"))
        hdr = "YES";
      else 
        hdr = "NO";
      
      if ( hdr.Equals("YES") && !DataLib.DataUtil.TableExists(schema, tableName, sqlConnStr) )
      {
        DataLib.DataUtil.CreateTableFromExcel(schema, tableName, sqlConnStr, excelFile, worksheet, hdr);
      }
      if (truncateOrAppend.ToUpper().Equals("TRUNCATE"))
      {
        DataLib.DataUtil.TruncateTable(sqlConnStr, schema, tableName);
      }
      
      Console.WriteLine("Load table = {0} from file = {1}", tableName, excelFile);
     
      DataLib.DataUtil.Excel2TableAddAuditCols(excelFile, worksheet, hdr, schema, tableName, sqlConnStr);
    }
    else if ( action.ToUpper().Equals(ACTION_EXCEL_FILES_IN_DIR2TABLE.ToUpper()) )
    {
      if (args.Length != 9)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String directory        = args[1];
      String filePattern      = args[2];
      String worksheet        = args[3];
      String header           = args[4];
      String schema           = args[5];
      String tableName        = args[6];
      String truncateOrAppend = args[7];
      String sqlConnStr       = args[8];
      
      String hdr = null;
      
      if (header.ToUpper().Equals("T"))
        hdr = "YES";
      else 
        hdr = "NO";
      
      if (truncateOrAppend.ToUpper().Equals("TRUNCATE"))
      {
        DataLib.DataUtil.TruncateTable(sqlConnStr, schema, tableName);
      }
     
      DataLib.DataUtil.ExcelFilesInDir2Table(directory, filePattern, worksheet, hdr, schema, tableName, sqlConnStr);
    }
    else if ( action.ToUpper().Equals(ACTION_FILES_IN_DIR2TABLE.ToUpper()) )
    {
      if (args.Length != 8)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String directory        = args[1];
      String filePattern      = args[2];
      String header           = args[3];
      String schema           = args[4];
      String tableName        = args[5];
      String truncateOrAppend = args[6];
      String sqlConnStr       = args[7];
      
      String hdr = null;
      
      if (header.ToUpper().Equals("T"))
        hdr = "YES";
      else 
        hdr = "NO";
      
      if (truncateOrAppend.ToUpper().Equals("TRUNCATE"))
      {
        DataLib.DataUtil.TruncateTable(sqlConnStr, schema, tableName);
      }
     
      DataLib.DataUtil.FilesInDir2Table(directory, filePattern, hdr, schema, tableName, sqlConnStr);
    }
    else if ( action.ToUpper().Equals(ACTION_GEN_EXCEL_FROM_MSSQL.ToUpper()) )
    {
      if (args.Length != 4)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String connectStr       = args[1];
      String selectStmt       = args[2];
      String excelFile        = args[3];
           
      Console.WriteLine("Calling GenExcelFromMSSQL");
      Console.WriteLine("  connectStr = {0}", connectStr);
      Console.WriteLine("  selectStmt = {0}", selectStmt);
      Console.WriteLine("  excelFile  = {0}", excelFile);
      
      DataLib.DataUtil.GenExcelFromMSSQL(connectStr, selectStmt, excelFile);
    }
    else if ( action.ToUpper().Equals(ACTION_GEN_EXCEL_FROM_MSSQL_QUERY.ToUpper()) )
    {
      if (args.Length != 4)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      String connectStr       = args[1];
      String queryFile        = args[2];
      String excelFile        = args[3];
      
      Console.WriteLine("  queryFile  = {0}", queryFile);
      String selectStmt = DataLib.DataUtil.FileToString(queryFile);
      
      Console.WriteLine("Calling GenExcelFromMSSQL");
      Console.WriteLine("  connectStr = {0}", connectStr);
      Console.WriteLine("  selectStmt = {0}", selectStmt);
      Console.WriteLine("  excelFile  = {0}", excelFile);
                  
      DataLib.DataUtil.GenExcelFromMSSQL(connectStr, selectStmt, excelFile);
    }
    else
    {
      Console.WriteLine("Unknown action = {0}\n", action);
      return;
    }      
      
  } // Main()

} // DataTools class
