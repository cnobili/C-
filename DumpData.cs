/*
 * DumpData.cs
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

namespace GetData
{

class DumpData
{

  /*
   * Private Data
   */
  private String dbOdbcDsn;
  private String dbUser;
  private String dbPass;
  private String dbSqlFile;
  private String delim;
  private String outputFile;
  private Boolean columnHeader;
  private OdbcConnection dbConn;
  private String dbSqlStmt;

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
    dbPass     = ConfigurationManager.AppSettings["dbPass"];
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
    Console.WriteLine("Disconnect from database");
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

  /*
   * Main() - Entry point of program.
   */
  public static void Main (String [] args)
  {
    DumpData dmp = new DumpData();

    try
    {
      Console.WriteLine("Starting Program DumpData");
      Console.WriteLine("Using configuration file = DumpData.exe.config");

      dmp.connectToDb();
      dmp.genFile();

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

} // DumpData namespace
