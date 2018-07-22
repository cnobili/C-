/*
 * GenSqlObj.cs
 *
 * Generates the source code for views, procedures, and functions from a
 * SQL Server database, using the passed in configuration file.
 *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using CodeLibrary;

namespace GenSqlObj
{

class GenSqlObj
{

  /*
   * Private Data
   */
  private static String dbServer;
  private static String dbUser;
  private static String dbPass;
  private static String dbCatalog;
  private static String dbSchema;
  private static String outputDir;

  /*
   * Public Methods
   */

  /*
   * Constructor.
   *
   * Gets database connection information from config file.
   */
  public GenSqlObj()
  {
    dbServer   = ConfigurationManager.AppSettings["dbServer"];
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
    
    dbCatalog  = ConfigurationManager.AppSettings["dbCatalog"];
    dbSchema   = ConfigurationManager.AppSettings["dbSchema"];
    outputDir  = ConfigurationManager.AppSettings["outputDir"];

  } // GenSqlObj()

  /*
   * GenObjs() - Generates DDL for tables, views, procedures, functions.
   */
  public void GenObjs(SqlConnection connection)
  {
		int n = 0;
		StreamWriter pw;
		String filename;
		String objCode;
		String objType;

    string queryStr = "select " +
                      "  o.name " +
                      ", m.definition " +
                      ", o.type " +
                      "from " +
                      "  sys.all_objects o " +
                      "  inner join " +
                      "  sys.all_sql_modules m " +
                      "  on m.object_id = o.object_id " +
                      "where 1 = 1 " +
                      "  and o.type in ('U', 'V', 'P', 'PC', 'FN', 'IF', 'SN') " +
                      "  and o.schema_id = SCHEMA_ID(@schema) " +
                      "order by " +
                      "  o.type ";

    string paramSchema = dbSchema;

    SqlCommand command = new SqlCommand(queryStr, connection);
    command.Parameters.AddWithValue("@schema", paramSchema);

    SqlDataReader reader = command.ExecuteReader();

    Console.WriteLine("\nConnected to Server = {0}, database = {1}", dbServer, dbCatalog);
    //Console.WriteLine("\nGenerate SQL code for procedures in schema = {0}", dbSchema);

    while (reader.Read())
    {
      filename = reader[0] + ".sql";
      objCode  = reader[1].ToString();
      objType  = reader[2].ToString();
      Console.WriteLine("  Writing out object type {0} to file {1}", objType, filename);
      n++;
      pw = new StreamWriter(outputDir + filename);
      pw.WriteLine(objCode);
      pw.Close();
    }

    reader.Close();
    Console.WriteLine("\nTotal Objects = {0}\n", n);

	} // GenObjs()

  public void GenProcs(SqlConnection connection)
  {
		int n = 0;
		StreamWriter pw;
		String filename;
		String objCode;

    string queryStr = "select " +
                      "  i.routine_name " +
                      ", object_definition(object_id(i.routine_name)) as obj_code " +
                      "from " +
                      "  information_schema.routines i " +
                      "where 1 = 1 " +
                      "  and i.routine_catalog = @database " +
                      "  and i.routine_schema = @schema " +
                      "order by " +
                      "  routine_type ";

    string paramCatalog = dbCatalog;
    string paramSchema  = dbSchema;

    SqlCommand command = new SqlCommand(queryStr, connection);
    command.Parameters.AddWithValue("@schema", paramSchema);
    command.Parameters.AddWithValue("@database", paramCatalog);

    SqlDataReader reader = command.ExecuteReader();

    Console.WriteLine("\nConnected to Server = {0}, database = {1}", dbServer, dbCatalog);
    Console.WriteLine("\nGenerate SQL code for procedures and functions in schema = {0}", dbSchema);

    while (reader.Read())
    {
      filename = reader[0] + ".sql";
      objCode  = reader[1].ToString();
      Console.WriteLine("  Writing out to file {0}", filename);
      n++;
      pw = new StreamWriter(outputDir + filename);
      pw.WriteLine(objCode);
      pw.Close();
    }

    reader.Close();
    Console.WriteLine("\nTotal Procedures and Functions = {0}\n", n);

	} // GenProcs()

  public void GenViews(SqlConnection connection)
  {
		int n = 0;
		StreamWriter pw;
		String filename;
		String objCode;

    string queryStr = "select " +
                      "  i.table_name " +
                      ", object_definition(object_id(i.table_name)) as obj_code " +
                      "from " +
                      "  information_schema.views i " +
                      "where 1 = 1 " +
                      "  and i.table_catalog = @database " +
                      "  and i.table_schema = @schema ";

    string paramCatalog = dbCatalog;
    string paramSchema  = dbSchema;

    SqlCommand command = new SqlCommand(queryStr, connection);
    command.Parameters.AddWithValue("@schema", paramSchema);
    command.Parameters.AddWithValue("@database", paramCatalog);

    SqlDataReader reader = command.ExecuteReader();

    Console.WriteLine("\nConnected to Server = {0}, database = {1}", dbServer, dbCatalog);
    Console.WriteLine("\nGenerate SQL code for views in schema = {0}", dbSchema);

    while (reader.Read())
    {
      filename = reader[0] + ".sql";
      objCode  = reader[1].ToString();
      Console.WriteLine("  Writing out to file {0}", filename);
      n++;
      pw = new StreamWriter(outputDir + filename);
      pw.WriteLine(objCode);
      pw.Close();
    }

    reader.Close();
    Console.WriteLine("\nTotal Views = {0}\n", n);

	} // GenViews()

  /*
   * Main() - Entry point of program.
   */
  public static void Main (String [] args)
  {
    // Get filename on the command line
    if ( !(args.Length == 1 || args.Length == 2) )
    {
      Console.WriteLine("Usage:GenSqlObj configFile [password]");
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
    GenSqlObj dmp = new GenSqlObj();

    //String connStr = "Data Source=" + dbServer + ";Initial Catalog=" + dbCatalog +
    //  ";Integrated Security=true";
    String connStr = "user id=" + dbUser + ";password=" + dbPass + ";server=" + dbServer + ";database=" + dbCatalog;

    using (SqlConnection connection = new SqlConnection(connStr))
    {
      connection.Open();
      //dmp.GenObjs(connection);
      dmp.GenProcs(connection);
      dmp.GenViews(connection);
      connection.Close();

      Console.WriteLine("Done!");
    }

  } // Main()

} // GenSqlObj class

} // GenSqlObj namespace

