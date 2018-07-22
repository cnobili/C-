/*
 * DumpData2.cs
 *
 * Dumps out a delimited flat file of data
 * for query referenced in configuration file.
 *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Text.RegularExpressions;
using CodeLibrary;

namespace DumpData2
{

class DumpData
{

  /*
   * Private Data
   */
  public const int    MAX_SQL_COL_NAME_LEN = 60;
  public const int    MAX_ORA_COL_NAME_LEN = 30;
  public const String msqlBulkInsPrefix    = "bulkIns_";
  public const String sqlSuffix            = ".sql";
  
  private static String dbOdbcDsn;
  private static String dbUser;
  private static String dbPass;
  private String dbSqlFile;
  private String delim;
  private String outputFile;
  private Boolean columnHeader;
  private OdbcConnection dbConn;
  private String dbSqlStmt;
  
  private String outputDir;
  private String oraExtTableName;
  private String oraExtDirectoryName;
  private String oraExtAddRecNum;
  private String msqlTableName;

  /*
   * Public Methods
   */

  /*
   * Constructor.
   *
   * Gets database connection information from config file.
   */
  public DumpData()
  {
    String colHeader;

    dbOdbcDsn  = ConfigurationManager.AppSettings["dbOdbcDsn"];
    dbUser     = ConfigurationManager.AppSettings["dbUser"];

    if (dbPass == null)
    {
      dbPass     = ConfigurationManager.AppSettings["dbPass"];
      if (dbPass.Equals(""))
      {
        Console.Write("Enter dbUser password:");
        dbPass = Util.ReadPassword();
      }
      else
      {
        dbPass = Util.Decrypt(dbPass);
      }
    }
    
    dbSqlFile  = ConfigurationManager.AppSettings["dbSqlFile"];
    delim      = ConfigurationManager.AppSettings["delim"];
    outputFile = ConfigurationManager.AppSettings["outputFile"];
    colHeader  = ConfigurationManager.AppSettings["columnHeader"];

    if (colHeader.Equals("true") )
      columnHeader = true;
    else
      columnHeader = false;

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
    
    outputDir = ConfigurationManager.AppSettings["outputDir"];
    oraExtTableName = ConfigurationManager.AppSettings["oraExtTableName"];
    oraExtDirectoryName = ConfigurationManager.AppSettings["oraExtDirectoryName"];
    oraExtAddRecNum = ConfigurationManager.AppSettings["oraExtAddRecNum"];
    msqlTableName = ConfigurationManager.AppSettings["msqlTableName"];

  } // DumpData()

  /*
   * connectToDb()
   *
   * Opens up a connection to the database.
   */
  public void connectToDb()
  {
    Console.WriteLine("\nMake connection to database URL = " + dbOdbcDsn);

    string connStr = "DSN=" + dbOdbcDsn + ";uid=" + dbUser + ";pwd=" + dbPass + ";";

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

  } // connectToDb()

  /*
   *  disconnectFromDb()
   *
   * Disconnects from the database.
   */
  public void disconnectFromDb()
  {
    Console.WriteLine("\nDisconnect from database");
    dbConn.Close();
    dbConn.Dispose();

  } // disconnectFromDb()
  
  /*
   * genFile()
   */
  public void genFile()
  {
    int recs = 0;
    StreamWriter pw = null;
    OdbcCommand cmdSQL = new OdbcCommand(dbSqlStmt, dbConn);
    OdbcDataReader dataReader = cmdSQL.ExecuteReader();
    int fieldCount = dataReader.FieldCount;

    String rec = null;
    String fieldSep = null;
    String col = null;
    int displayCount = 100000;

    Console.WriteLine("Execute SQL = " + dbSqlStmt);
    Console.WriteLine("Number of columns in select stmt = " + fieldCount);

    try
    {
      pw = new StreamWriter(outputFile);

      if (columnHeader)
      {
        rec = "";
        fieldSep = "";
        for (int i = 0; i < fieldCount; i++)
        {
          rec += fieldSep;
          rec += dataReader.GetName(i);
          fieldSep = delim;
        }
        pw.WriteLine(rec);
      }

      // Iterate through the resultset
      while (dataReader.Read())
      {
        rec = "";
        fieldSep = "";
        for (int i = 0; i < fieldCount; i++)
        {
          rec += fieldSep;
          col = dataReader[i].ToString();
          if (col != null)
            rec += col;
          else
            rec += "";
          fieldSep = delim;
        }

        recs++;

        if (recs % displayCount == 0)
          Console.WriteLine("Total records written so far = " + recs);
        pw.WriteLine(rec);
      }

      Console.WriteLine("\nTotal records written out to file = " + recs);

    }
    finally
    {
      pw.Close();
    }

  } // genFile()
  
  public static String PadRight(String s, int n)
  {
    return String.Format("{0," + -n + "}", s);

  } // PadRight()

  public void GenTsqlTable()
  {
    String colDef;
    FileInfo fi = new FileInfo(outputFile);
    StreamReader sr = fi.OpenText();
    String header = sr.ReadLine();
    sr.Close();

    String [] tokens = header.Split(new Char [] {System.Convert.ToChar(delim)});

    StreamWriter pw = new StreamWriter(outputDir + "\\" + msqlTableName + sqlSuffix);

    pw.WriteLine("create table " + msqlTableName);
    pw.WriteLine("(");

    for (int i = 0; i < tokens.Length; i++)
    {
      colDef = tokens[i].ToLower();
      if (colDef.Length > MAX_SQL_COL_NAME_LEN)
      {
        colDef.Substring(0, MAX_SQL_COL_NAME_LEN);
      }

      colDef = PadRight(colDef, MAX_SQL_COL_NAME_LEN) + " varchar(8000)";

      if (i == 0)
      {
        pw.WriteLine("  " + colDef);
      }
      else
      {
        pw.WriteLine(", " + colDef);
      }
    }

    pw.WriteLine(")");

    pw.Close();

  } // GenTsqlTable()

  public void GenTsqlBulkIns()
  {
    StreamWriter pw = new StreamWriter(outputDir + "\\" + msqlBulkInsPrefix + msqlTableName + sqlSuffix);

    pw.WriteLine("bulk insert " + msqlTableName);
    pw.WriteLine("from '" + outputFile + "'");
    pw.WriteLine("with");
    pw.WriteLine("(");
    pw.WriteLine("  datafiletype = 'char'");
    pw.WriteLine(", firstrow = 2");
    pw.WriteLine(", fieldterminator = '" + delim + "'");
    pw.WriteLine(", errorfile = '" + outputDir + "\\\\" + msqlTableName + ".err'");
    pw.WriteLine(")");

    pw.Close();

  } // GenTsqlBulkIns()
  
  public void GenOraExtTab()
  {
    String filename = Path.GetFileName(outputFile);
    
    String colDef;  
    FileInfo fi = new FileInfo(outputFile);
    StreamReader sr = fi.OpenText();
    String header = sr.ReadLine();
    sr.Close();

    String [] tokens = header.Split(new Char [] {System.Convert.ToChar(delim)});
  
    StreamWriter pw = new StreamWriter(outputDir + "\\" + oraExtTableName + sqlSuffix);
    
    pw.WriteLine("create table " + oraExtTableName);
    pw.WriteLine("(");

    for (int i = 0; i < tokens.Length; i++)
    {
      if (tokens[i].Length == 0) continue;

      colDef = tokens[i].ToLower();
      colDef = Regex.Replace(colDef, @"[- ,']", "_");

      if (colDef.Length > MAX_ORA_COL_NAME_LEN)
      {
        colDef = colDef.Substring(0, MAX_ORA_COL_NAME_LEN);
      }

      colDef = PadRight(colDef, MAX_ORA_COL_NAME_LEN) + " varchar(4000)";

      if (i == 0)
      {
        pw.WriteLine("  " + colDef);
      }
      else
      {
        pw.WriteLine(", " + colDef);
      }
    }

    if (oraExtAddRecNum.Equals("Y"))
    {
      colDef = PadRight("rec_num", MAX_ORA_COL_NAME_LEN) + " number";
      pw.WriteLine(", " + colDef);
    }

    pw.WriteLine(")");
    pw.WriteLine("organization external");
    pw.WriteLine("(");
    pw.WriteLine("  type oracle_loader");
    pw.WriteLine("  default directory " + oraExtDirectoryName.ToUpper());
    pw.WriteLine("  access parameters");
    pw.WriteLine("  (");
    pw.WriteLine("    records delimited by NEWLINE");
    pw.WriteLine("    readsize 5000000");
    pw.WriteLine("    skip 1");
    pw.WriteLine("    nobadfile");
    pw.WriteLine("    nodiscardfile");
    pw.WriteLine("    nologfile");
    pw.WriteLine("    fields terminated by '" + delim + "'");
    pw.WriteLine("    missing field values are null");
    pw.WriteLine("    reject rows with all null fields");

    if (oraExtAddRecNum .Equals("Y"))
    {
      pw.WriteLine("    (");

      for (int i = 0; i < tokens.Length; i++)
      {
        if (tokens[i].Length == 0) continue;

        colDef = tokens[i].ToLower();
        Regex.Replace(colDef, @"[- ,']", "_");

        if (colDef.Length > MAX_ORA_COL_NAME_LEN)
        {
          colDef = colDef.Substring(0, MAX_ORA_COL_NAME_LEN);
        }

        colDef = PadRight(colDef, MAX_ORA_COL_NAME_LEN);

        if (i == 0)
        {
          pw.WriteLine("      " + colDef);
        }
        else
        {
          pw.WriteLine("    , " + colDef);
        }
      }

      colDef = PadRight("rec_num", MAX_ORA_COL_NAME_LEN) + " recnum";
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

  } // GenOraExtTab()

  /*
   * Main() - Entry point of program.
   */
  public static void Main (String [] args)
  {
    // Get filename on the command line
	if ( !(args.Length == 1 || args.Length == 2) )
	{
	  Console.WriteLine("Usage:DumpData2 configFile [password]");
    Console.WriteLine("");
    Console.WriteLine("  configFile - configuration file");
    Console.WriteLine("  password   - optional cleartext password");
	  return;
	}
	String configFile = args[0];
  if (args.Length == 2) dbPass = args[1];

	if (File.Exists(configFile))
	{
	  Console.WriteLine("Using Configuration file = <{0}>", configFile);
	}
	else
	{
	  Console.WriteLine("filename = <{0}> not found", configFile);
	  return;
	}

    AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", configFile);
    DumpData dmp = new DumpData();

    try
    {
      dmp.connectToDb();
      dmp.genFile();
      
      if (dmp.msqlTableName != null)
      {
        Console.WriteLine();
        Console.WriteLine("Writing out TSQL Table DDL script   = {0}{1}", dmp.msqlTableName, sqlSuffix);
        Console.WriteLine("Writing out TSQL Bulk Insert script = {0}{1}{2}", msqlBulkInsPrefix, dmp.msqlTableName, sqlSuffix);
        dmp.GenTsqlTable();
        dmp.GenTsqlBulkIns();
      }
      if (dmp.oraExtTableName != null)
      {
        Console.WriteLine();
        Console.WriteLine("Writing out Oracle External Table script = {0}{1}", dmp.oraExtTableName, sqlSuffix);
        dmp.GenOraExtTab();
      }
    
    }
    catch(Exception e)
    {
      Console.WriteLine(e.ToString());
    }
    finally
    {
      dmp.disconnectFromDb();
      Console.WriteLine("Done!");
    }

  } // Main()

} // DumpData class

} // DumpData2 namespace
