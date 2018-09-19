/*
 * DataLib.cs
 *
 * Library of data methods for databases and flat files.
 *
 * Craig Nobili
 */
 
using System;
using System.Text;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Collections;
using System.Collections.Generic;


namespace DataLib
{

  public class DataUtil
  {

    /*
     * Public Members
     */
     
    public const String TAB                  = "\t";
    public const String FILE_DATA_TYPE       = "SQLCHAR";
    public const String SQL_SUFFIX           = ".sql";
    public const String MSQL_FILE_SEL_PREFIX = "fileSel_";

    /*
     * Private Methods
     */
     
    /*
     * Public Methods
     */

    public static void GenSchemaINI(String fullPathSchemaINI, String filename, String delim, String colNameHeader)
    {
      StreamWriter sw = new StreamWriter(fullPathSchemaINI);
      sw.WriteLine("[{0}]", filename);
      sw.WriteLine("ColNameHeader={0}", colNameHeader);
      sw.WriteLine("Format=Delimited({0})", delim);
      sw.Close();
    
    } // GenSchemaINI()
    
    public static string CleanString( string s )
    {
        if (s != null && s.Length > 0)
        {
            StringBuilder sb = new StringBuilder( s.Length );
            foreach (char c in s)
            {
                sb.Append( Char.IsControl( c ) ? ' ' : c );
            }
            s = sb.ToString();
        }
        return s;
        
    } // CleanString()

    public static String StripNonPrintableCharsFromString(string inputValue, string replaceStr = null)
    {
        if (replaceStr == null)
        {
            return Regex.Replace(inputValue, @"[^\u0000-\u001F]", String.Empty);
        }
        else
        {
            return Regex.Replace(inputValue, @"[^\u0000-\u001F]", replaceStr);
        }
    }

    public static String PadRight(String s, int n)
    {
      return String.Format("{0," + -n + "}", s);

    } // PadRight()
    
    public static void TruncateTable(String dbConnectStr, String schema, String table) 
    {
      String sqlStmt = "truncate table " + schema + "." + table;
      SqlCommand cmd = new SqlCommand();
      SqlConnection conn = new SqlConnection(dbConnectStr);
        
      conn.Open();
      cmd.Connection = conn;
      cmd.CommandText = sqlStmt;
      cmd.ExecuteNonQuery();
      conn.Close();

    } // TruncateTable()

  public static Boolean TableExists(String schema, String tableName, String connectString)
  {
  
    var connection = new SqlConnection(connectString);
    String sqlSelCount = "select count(*) from INFORMATION_SCHEMA.TABLES where TABLE_SCHEMA = '" + schema + "' and table_name = '" + tableName + "'";
    var command = new SqlCommand(connectString);
    
    connection.Open();
    command = new SqlCommand(sqlSelCount, connection);
    int numRecs = (int)command.ExecuteScalar();
  
    connection.Close();
    
    if (numRecs == 0)
      return(false);
    return(true);
     
  } // TableExists()
  
  public static void CreateTable(String schema, String tableName, String connectString, string filename, string delim)
  {
  
    // Create table in destination sql database to hold file data
    var sql = new StringBuilder();
    string[] columns = null;
    
    if (!File.Exists(filename))
    {
      Console.WriteLine("File = {0} does not exist", filename);
      return;
    }
    
    using (var stream = new StreamReader(filename))
    {
      var fieldNames = stream.ReadLine();
      //if (fieldNames != null) columns = fieldNames.Split("\t".ToCharArray());
      if (fieldNames != null) columns = fieldNames.Split(delim.ToCharArray());
    }
  
    sql.Append("CREATE TABLE ");
    sql.Append(schema + "." + tableName);
    sql.Append(" ( ");
      
    foreach (var columnName in columns)
    {
      if (columns.GetUpperBound(0) == Array.IndexOf(columns, columnName))
      {
        sql.Append(columnName);
        sql.Append(" VARCHAR(8000))");
      }
      else
      {
        sql.Append(columnName);
        sql.Append(" VARCHAR(8000),");
      }
    }
    
    // remove last 2 characters
    sql.Length--;
    sql.Length--;
    sql.Append(")"); // put back the closing parenthesis on the data type
    sql.Append(")"); // final closing parenthesis for create table statemetn
     
    var connection = new SqlConnection(connectString);
    var command = new SqlCommand(sql.ToString(), connection);
  
    connection.Open();
    try 
    {
      command.ExecuteNonQuery();
    }
    catch (Exception ex)
    {
      Console.WriteLine(ex.ToString());
    }
    finally
    {
      connection.Close();
    }
  
  } // CreateTable()

    public static void CreateTableFromExcel(String schema, String tableName, String connectString, string excelFile, string excelWorksheet, string hdr)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=" + hdr + ";IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=" + hdr + ";IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, tableName);
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      Console.WriteLine("\nThe columns are:");

      foreach (DataColumn col in item.Columns)
      {
        // Remove any generated headers even though HDR=YES
        if (col.ColumnName.ToString().Substring(0,2).Equals("F1") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F2") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F3") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F4") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F5") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F6") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F7") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F8") ||
            col.ColumnName.ToString().Substring(0,2).Equals("F9")
           )
        {
      Console.WriteLine("Skipping column {0}", col.ColumnName);
      }
      else
      {
          colName.Add(col.ColumnName);
          colType.Add(col.DataType);
          Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      }
      Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate create table ...\n");

      String ddl = "Create table " + tableName + "\n" + "( \n";

      for (int i =0; i < colName.Count; i++)
      {
        if (i == 0)
          ddl = ddl + "  " + colName[i];
        else
          ddl = ddl + ", " + colName[i];

        //if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
        //  ddl = ddl + " decimal(20,4) \n";
        //else
          ddl = ddl + " varchar(8000) \n";
      }
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      var connection = new SqlConnection(connectString);
      var command2 = new SqlCommand(ddl, connection);
  
      connection.Open();
      try 
      {
        command2.ExecuteNonQuery();
      }
      catch (Exception ex)
      {
        Console.WriteLine(ex.ToString());
      }
      finally
      {
        connection.Close();
      }

    } // CreateTableFromExcel()
    
    //
    // Runs an OS command
    //
    public static void RunCmdSync(string command)
    {
      //System.Diagnostics.Process.Start(@"C:\MyScript.bat");

      // create the ProcessStartInfo using "cmd" as the program to be run,
      // /c tells cmd that we want it to execute the command that follows, and then exit.
      System.Diagnostics.ProcessStartInfo procStartInfo =
        new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command);

      // The following commands are needed to redirect the standard output.
      procStartInfo.RedirectStandardOutput = true;
      procStartInfo.UseShellExecute = false;
      // Do not create the black window.
      procStartInfo.CreateNoWindow = true;
      // Now we create a process, assign its ProcessStartInfo and start it
      System.Diagnostics.Process proc = new System.Diagnostics.Process();
      proc.StartInfo = procStartInfo;
      proc.Start();
      // Get the output into a string
      string result = proc.StandardOutput.ReadToEnd();
      // Display the command output.
      //Console.WriteLine(result);

    } // RunCmdSync()

    public static void Msg(String s)
    {
      Console.WriteLine("{0} => {1}", DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss"), s);

    } // Msg()
    public static bool FileExists(String path)
    {
      return(File.Exists(path));
    
    } // FileExists()
    
    public static String[] ReadFileIntoArray(String filePath)
    {
      if ( !FileExists(filePath) ) return(new String[]{});
        
      String[] lines = System.IO.File.ReadAllLines(filePath);
      //String[] lines = theText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
    
      return(lines);
  
    } // ReadFileIntoArray()
    
    public static long CountLinesInFile(String filename)
    {
      String[] lines = File.ReadAllLines(filename);
      return(lines.Length);
    
    } // CountLinesInFile()
    
    public static void AddHeader2File(String filename, String delimiter, String newFilename)
    {
      String[] lines = File.ReadAllLines(filename);
      String[] tokens = lines[0].Split(delimiter.ToCharArray());
      String header = null;
      String delim = null;
    
      for (int i = 0; i < tokens.Length; i++)
      {
        header = header + delim + "column" + (i + 1).ToString("D2");
        delim = delimiter;
      }
    
      // File.WriteAllLines(path, createText, Encoding.UTF8);
    
      // Need header first, use LINQ
      List<string> list = new List<string>(lines);
      StreamWriter sw = new StreamWriter(newFilename);
      sw.WriteLine(header);
      list.ForEach(r=> sw.WriteLine(r));
      sw.Close();
    
    } // AddHeader2File()
    
    public static void GenTsqlFormatFile(String dataFile, String colDelim, String rowDelim)
    {
      String fileDir = Path.GetDirectoryName(dataFile);
      String fileNameNoExt = Path.GetFileNameWithoutExtension(dataFile);
      FileInfo fi = new FileInfo(dataFile);
      StreamReader sr = fi.OpenText();
      String header = sr.ReadLine();
      sr.Close();

      String [] tokens = header.Split(new Char [] {System.Convert.ToChar(colDelim)});

      StreamWriter pw = new StreamWriter(fileDir + "\\" + fileNameNoExt + ".fmt");

      pw.WriteLine(tokens.Length); // number of columns

      for (int i = 0; i < tokens.Length; i++)
      {
        pw.Write("{0}{1}", i + 1, TAB);          // field order in host file
        pw.Write("{0}{1}", FILE_DATA_TYPE, TAB); // host file data type
        pw.Write("{0}{1}", 0, TAB);              // prefix length
        pw.Write("{0}{1}", 0, TAB);              // host file data length (use zero)
        if (i == tokens.Length - 1)
          pw.Write("{0}{1}", rowDelim, TAB);     // last column delimiter (i.e. row delimiter)
        else
          pw.Write("{0}{1}", colDelim, TAB);     // column delimiter
        pw.Write("{0}{1}", i + 1, TAB);          // server column order
        pw.Write("{0}{1}", tokens[i], TAB);      // server column name
        pw.Write("{0}", "\"\"");                 // column collation
        pw.WriteLine();
      }

      pw.Close();

    } // GenTsqlFormatFile()
    
    public static void GenTsqlReadFileAsTable(String dataFile, String colDelim, String rowDelim)
    {
      String outputDir = Path.GetDirectoryName(dataFile);
      String filename = Path.GetFileName(dataFile);
      String filenameNoExt = Path.GetFileNameWithoutExtension(dataFile);

      StreamWriter pw = new StreamWriter(outputDir + "\\" + MSQL_FILE_SEL_PREFIX + filenameNoExt + SQL_SUFFIX);
      Console.WriteLine("Writing out file as table T-SQL statement to file = {0}", outputDir + "\\" + MSQL_FILE_SEL_PREFIX + filenameNoExt + SQL_SUFFIX);

      pw.WriteLine("select *");
      pw.WriteLine("from openrowset");
      pw.WriteLine("(");
      pw.WriteLine("  BULK '{0}'", dataFile);
      pw.WriteLine(", formatfile = '{0}\\{1}.fmt'", outputDir, filenameNoExt);  
      pw.WriteLine(", firstrow = 2");
      pw.WriteLine(", errorfile = '{0}\\{1}.err'", outputDir, filenameNoExt);
      pw.WriteLine(") as {0}", filenameNoExt);

      pw.Close();

      // Generate format file
      GenTsqlFormatFile(dataFile, colDelim, rowDelim);

    } // GenTsqlReadFileAsTable()
    
    public static void GenSelectStmt(String outputfile, String schemaName, String tableName, String server, String database, String user , String pass)
    {
      String dbConnectStr;
      StreamWriter pw = null;
      SqlConnection dbConn = null;
      String query;

      if (user == null)
      {
        dbConnectStr = "Server=" + server + ";Database=" + database + ";Trusted_Connection=True";
      }
      else
      {
        dbConnectStr = "Server=" + server + ";Database=" + database + ";user id=" + user + ";password=" + pass;
      }

      query = "select top 1 * from " + schemaName + "." + tableName;     
      try
      {
        dbConn = new SqlConnection(dbConnectStr);
        dbConn.Open();
        Console.WriteLine("Opened database connection");
        SqlCommand cmdSQL = new SqlCommand(query, dbConn);
        SqlDataReader dataReader = cmdSQL.ExecuteReader();
        int fieldCount = dataReader.FieldCount;

        String fieldSep = null;

        Console.WriteLine("Execute SQL = " + query);
        Console.WriteLine("Number of columns in select stmt = " + fieldCount);

        pw = new StreamWriter(outputfile);
        pw.WriteLine("select");

        fieldSep = "  ";
        for (int i = 0; i < fieldCount; i++)
        {
          pw.WriteLine(fieldSep + dataReader.GetName(i));
          fieldSep = ", ";
        }
        pw.WriteLine("from");
        pw.WriteLine("  " + schemaName + "." + tableName);
        pw.WriteLine(";");
      }
      catch(Exception e)
      {
        Console.WriteLine(e.ToString());
      }
      finally
      {
        pw.Close();
        dbConn.Close();
      }

    } // GenSelectStmt()

    public static void OraGenExtTableDDL(String filePath, String filename, String colDelimiter, String recDelimiter, 
                                   String table, String dirObj, String addRecNum, String outputDir)
    {
      String header;
      String colDef;
      FileInfo fi = new FileInfo(filePath);
      StreamReader sr = fi.OpenText();

      // The end of file character can be something other than carriage return/line feeed
      // We just need to get the first record which is the column header
      // Assume header record is less than 16K bytes, read this many or less bytes
      // in order to split out the header record to get the column names
      char [] buf = new char[16384];
      sr.Read(buf, 0, 16384);
      sr.Close();

      header = new String(buf);
      String [] tok = Regex.Split(header, recDelimiter);
      header = tok[0];
      String [] tokens = header.Split(colDelimiter.ToCharArray());

      StreamWriter pw = new StreamWriter(outputDir + "\\" + table + ".sql");
      Console.WriteLine("Writing out file = <{0}>", outputDir + "\\" + table + ".sql");

      pw.WriteLine("create table " + table);
      pw.WriteLine("(");

      for (int i = 0; i < tokens.Length; i++)
      {
        if (tokens[i].Length == 0) continue;

        colDef = tokens[i].ToLower();
        colDef = Regex.Replace(colDef, @"[- ,']", "_");

        if (colDef.Length > 80)
        {
          colDef = colDef.Substring(0, 80);
        }

        colDef = PadRight(colDef, 80) + " varchar(4000)";

        if (i == 0)
        {
          pw.WriteLine("  " + colDef);
        }
        else
        {
          pw.WriteLine(", " + colDef);
        }
      }

      if (addRecNum.Equals("Y"))
      {
        colDef = PadRight("rec_num", 80) + " number";
        pw.WriteLine(", " + colDef);
      }

      pw.WriteLine(")");
      pw.WriteLine("organization external");
      pw.WriteLine("(");
      pw.WriteLine("  type oracle_loader");
      pw.WriteLine("  default directory " + dirObj.ToUpper());
      pw.WriteLine("  access parameters");
      pw.WriteLine("  (");
      pw.WriteLine("    records delimited by '" + recDelimiter + "'");
      pw.WriteLine("    readsize 5000000");
      pw.WriteLine("    skip 1");
      pw.WriteLine("    nobadfile");
      pw.WriteLine("    nodiscardfile");
      pw.WriteLine("    nologfile");
      pw.WriteLine("    fields terminated by '" + colDelimiter + "'");
      pw.WriteLine("    missing field values are null");
      pw.WriteLine("    reject rows with all null fields");

      if (addRecNum .Equals("Y"))
      {
        pw.WriteLine("    (");

        for (int i = 0; i < tokens.Length; i++)
        {
          if (tokens[i].Length == 0) continue;

          colDef = tokens[i].ToLower();
          Regex.Replace(colDef, @"[- ,']", "_");

          if (colDef.Length > 30)
          {
            colDef = colDef.Substring(0, 30);
          }

          colDef = PadRight(colDef, 30);

          if (i == 0)
          {
            pw.WriteLine("      " + colDef);
          }
          else
          {
            pw.WriteLine("    , " + colDef);
          }
        }

        colDef = PadRight("rec_num", 30) + " recnum";
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

    } // OraGenExtTableDDL()

    public static String BuildBcpCmd(String server, String database, String schema, String table, String datafile, String fieldDelimiter, String outputDir, int firstRow, long maxErrors, long batchSize, String user, String pass)
    {
      String cmd;
      String logfile = outputDir + "\\" + table + ".log";
      String errfile = outputDir + "\\" + table + ".err";
      String delim = "\"" + fieldDelimiter + "\""; // escape fieldDelimiter in case | is used as shell will treat as pipe redirection
    
      // bcp ClarityMonitor.dbo.bcp_test in bcp_test.dat -t "|" -F 2 -S REVDASHDBDEV -T -c -o bcp_test.log -e bcp_test.err -m 10 -b 1
    
      if (user == null)
      {
        cmd = "bcp " + database + "." + schema + "." + table + " in " + datafile + " -t " + delim + " -F " + firstRow + " -S " + server + " -T -c -o " + logfile + " -e " + errfile + " -m " + maxErrors + " -b " + batchSize;
      }
      else
      {
        cmd = "bcp " + database + "." + schema + "." + table + " in " + datafile + " -t " + delim + " -F " + firstRow + " -S " + server + " -U " +user + " -P " + pass + " -c -o " + logfile + " -e " + errfile + " -m " + maxErrors + " -b " + batchSize;
      }
    
      return(cmd);
        
    } // BuildBcpCmd()
    
  //
  // Extracts data and writes to flat file
  //
  public static void BcpExtractData(String outputfile, String delimiter, String query, String server, String database, String user , String pass)
  {
    String cmd;
 
    if (user == null || user == "")
    {
      cmd = "bcp \"" + query + "\" queryout " + outputfile + " -T -t\"" + delimiter + "\"" + " -S " + server + " -d " + database + " -c";
    }
    else
    {
      cmd = "bcp \"" + query + "\" queryout " + outputfile + " -U " + user + " -P " + pass + " -t\"" + delimiter + "\"" + " -S " + server + " -d " + database + " -c";
    }
    
    Console.WriteLine("Executing command-> {0}", cmd);
    RunCmdSync(cmd);
    
  } // BcpExtractData()
  
  //
  // Bulk Loads table from data file - Overloaded method
  //
  public static void BcpLoad(String server, String database, String schema, String table, String filePath, String delimiter, String outputDir, int firstRow, long maxErrors, long batchSize, String user , String pass)
  {
    String cmd;
    
    if (!File.Exists(filePath))
    {
      Console.WriteLine("File = {0} does not exist", filePath);
      return;
    }
    FileInfo fi = new FileInfo(filePath);
    if (fi.Length == 0)
    {
      Console.WriteLine("File = {0} is empty, zero bytes", filePath);
      return;
    }
    
    cmd = BuildBcpCmd(server, database, schema, table, filePath, delimiter, outputDir, firstRow, maxErrors, batchSize, user, pass);
    Console.WriteLine("Executing command:");
    Console.WriteLine(cmd);
    RunCmdSync(cmd);
    
  } // BcpLoad()
  
  //
  // Bulk Loads table from data file - Overloaded method
  //
  public static void BcpLoad(String server, String database, String schema, String table, String filePath, String delimiter, String outputDir, int firstRow, long maxErrors, long batchSize, String user , String pass, String appendOrTruncate)
  {
    String dbConnectStr;
        
    if (!File.Exists(filePath))
    {
      Console.WriteLine("File = {0} does not exist", filePath);
      return;
    }
    FileInfo fi = new FileInfo(filePath);
    if (fi.Length == 0)
    {
      Console.WriteLine("File = {0} is empty, zero bytes", filePath);
      return;
    }
    
    if (user == null || user == "")
    {
      dbConnectStr = "Server=" + server + ";Database=" + database + ";Trusted_Connection=True";
    }
    else
    {
      dbConnectStr = "Server=" + server + ";Database=" + database + ";Trusted_Connection=False" + "user id=" + user + ";password=" + pass;
    }
     
    if (appendOrTruncate.ToUpper().Equals("TRUNCATE"))
    {
      TruncateTable(dbConnectStr, schema, table);
    }
      
    BcpLoad(server, database, schema, table, filePath, delimiter, outputDir, firstRow, maxErrors, batchSize, user, pass);
     
  } // BcpLoad()
  
  public static void GenTableDDL(String filePath, String outputDir, String table, String rec_delimiter, String col_delimiter)
  {
    String colDef;
    FileInfo fi = new FileInfo(filePath);
    StreamReader sr = fi.OpenText();

    // The end of file character can be something other than carriage return/line feeed
    // We just need to get the first record which is the column header
    // Assume header record is less than 16K bytes, read this many or less bytes
    // in order to split out the header record to get the column names
    char [] buf = new char[16384];
    sr.Read(buf, 0, 16384);
    sr.Close();

    String header = new String(buf);
    String [] tok = Regex.Split(header, rec_delimiter);
    header = tok[0];
    String [] tokens = header.Split(col_delimiter.ToCharArray());

    StreamWriter pw = new StreamWriter(outputDir + "\\" + table + ".sql");
    Console.WriteLine("Writing out file = <{0}>", outputDir + "\\" + table + ".sql");

    pw.WriteLine("create table " + table);
    pw.WriteLine("(");

    for (int i = 0; i < tokens.Length; i++)
    {
      if (tokens[i].Length == 0) continue;

      colDef = tokens[i].ToLower();
      colDef = Regex.Replace(colDef, @"[- ]", "_");

      if (colDef.Length > 80)
      {
        colDef = colDef.Substring(0, 80);
      }

      colDef = PadRight(colDef, 80) + " varchar(8000)";

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

  } // GenTableDDL()
  
  public static void GenBulkInsertStmt(String outputDir, String table, String filePath, String col_delimiter, String rec_delimiter)
  {
    StreamWriter pw = new StreamWriter(outputDir + "\\load_" + table + ".sql");
    Console.WriteLine("Writing out file = <{0}>", outputDir + "\\load_" + table + ".sql");

    pw.WriteLine("bulk insert " + table);
    pw.WriteLine("from '" + filePath + "'");
    pw.WriteLine("with");
    pw.WriteLine("(");
    pw.WriteLine("  datafiletype = 'char'");
    pw.WriteLine(", firstrow = 2");
    pw.WriteLine(", fieldterminator = '" + col_delimiter + "'");
    pw.WriteLine(", rowterminator = '" + rec_delimiter + "'");
    pw.WriteLine(", errorfile = '" + outputDir + "\\" + table + ".err'");
    pw.WriteLine(")");

    pw.Close();

  } // GenBulkInsertStmt()
  
  // Generates a Bulk Insert Statement from a flat file, the flat file must have a header record
  public static void Csv2BulkInsert(String filePath, String col_delimiter, String rec_delimiter, String table, String outputDir)
  {
    GenTableDDL(filePath, outputDir, table, rec_delimiter, col_delimiter);
    GenBulkInsertStmt(outputDir, table, filePath, col_delimiter, rec_delimiter);
  
  } // Csv2BulkInsert()
  
  public static void BulkCopyTable(String srcConnStr, String srcTable, String dstConnStr, String dstTable, int batchSize)
  {

    using (SqlConnection sourceConnection = new SqlConnection(srcConnStr))
    {
      sourceConnection.Open();

      SqlCommand srcCommandRowCount = new SqlCommand
      (
        "select " +
        "  count(*) " +
        "from " +
          srcTable + " with (nolock) " 
      ,  sourceConnection
      );
      
      long dstBeforeTableRecs;
      long srcTableRecs = System.Convert.ToInt32(
      srcCommandRowCount.ExecuteScalar());
      Console.WriteLine("Source Table {0}, Records = {1}", srcTable, srcTableRecs);
      
      using (SqlConnection destinationConnection = new SqlConnection(dstConnStr))
      {
        destinationConnection.Open();
        
        SqlCommand dstCommandRowCount = new SqlCommand
        (
          "select " +
          "  count(*) " +
          "from " +
             dstTable + " with (nolock) " 
        ,  destinationConnection
        );
            
        dstBeforeTableRecs = System.Convert.ToInt32(
        dstCommandRowCount.ExecuteScalar());
        Console.WriteLine("Destination Table {0} Before Insert, Records = {1}", dstTable, dstBeforeTableRecs);
      }

      // Get data from the source table as a SqlDataReader.
      SqlCommand commandSourceData = new SqlCommand
      (
        "select * " +
        "from " +
          srcTable + "  with (nolock) " 
      , sourceConnection
      );
      
      SqlDataReader reader = commandSourceData.ExecuteReader();

      // Open the destination connection
      using (SqlConnection destinationConnection = new SqlConnection(dstConnStr))
      {
        destinationConnection.Open();
        
        // Set up the bulk copy object.  
        // Note that the column positions in the source 
        // data reader match the column positions in  
        // the destination table so there is no need to 
        // map columns. 
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
        {
          bulkCopy.DestinationTableName = dstTable;
          bulkCopy.BulkCopyTimeout = 0; // no timeout
          bulkCopy.BatchSize = batchSize; // 0 is default and commits entire result set

          try
          {
            // Write from the source to the destination.
            bulkCopy.WriteToServer(reader);
          }
          catch (Exception ex)
          {
            Console.WriteLine(ex.Message);
          }
          finally
          {
            // Close the SqlDataReader. The SqlBulkCopy 
            // object is automatically closed at the end of the using block.
            reader.Close();
          }
        }

        // Perform a final count on the destination table to see how many rows were added.
        
        SqlCommand commandRowCount = new SqlCommand();
        commandRowCount.Connection = destinationConnection;
        commandRowCount.CommandText = 
        "select " +
        "  count(*) " +
        "from " +
        dstTable + "  with (nolock) "
        ;
        long dstAfterTableRecs = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
        Console.WriteLine("Destination Table {0} After Insert, Records = {1}", dstTable, dstAfterTableRecs);
        Console.WriteLine("{0} rows were added.", dstAfterTableRecs - dstBeforeTableRecs);
        //Console.WriteLine("Press Enter to finish.");
        //Console.ReadLine();
      }
    }  
  
  } // BulkCopyTable()
  
  public static void OdbcQuery2SqlTable(String odbcConnStr, String query, String dstConnStr, String schemaName, String dstTable, Boolean truncateFlg)
  {
    OdbcDataAdapter da = new OdbcDataAdapter(query, odbcConnStr);
      
    DataTable dt = new DataTable(schemaName + "." + dstTable);
    da.Fill(dt);
      
    /*
    foreach (DataRow row in dt.Rows)
    {
      Console.WriteLine();
      for(int x = 0; x < dt.Columns.Count; x++)
      {
        Console.Write(row[x].ToString() + " ");
      }
    }
    */
    
    if (dt.Rows.Count == 0)
    {
      Console.WriteLine("Zero rows selected");
      return;
    }
    
    Console.WriteLine("{0} rows selected", dt.Rows.Count );
    
    if (truncateFlg)
    {
      TruncateTable(dstConnStr, schemaName, dstTable);
    }
      
    using (SqlConnection sqlConn = new SqlConnection(dstConnStr))
    {
      sqlConn.Open();
      using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
      {
        
        //foreach (DataColumn c in dt.Columns)
        //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
       
        bulkCopy.DestinationTableName = dt.TableName;
        try
        {
          bulkCopy.WriteToServer(dt);
        }
        catch (Exception ex)
        {
          Console.WriteLine(ex.Message);
        }
      }
    }
      
  } // OdbcQuery2SqlTable()
  
  /*
   * FileToString()
   */
  public static String FileToString(String filePath)
  {
    if (!File.Exists(filePath))
    {
      Console.WriteLine("Error: file = <{0}> does not exist\n", filePath);
      return null;
    }
    
    return File.ReadAllText(filePath);
  
  } // FileToString()
    
  /*
   * BulkLoadQuery2Table()
   */
  public static void BulkLoadQuery2Table(String srcConnStr, String srcQuery, String dstConnStr, String dstTable, int batchSize)
  {

    using (SqlConnection sourceConnection = new SqlConnection(srcConnStr))
    {
      sourceConnection.Open();
      
      long dstBeforeTableRecs;
      
      using (SqlConnection destinationConnection = new SqlConnection(dstConnStr))
      {
        destinationConnection.Open();
        
        SqlCommand dstCommandRowCount = new SqlCommand
        (
          "select " +
          "  count(*) " +
          "from " +
             dstTable + " with (nolock) " 
        ,  destinationConnection
        );
            
        dstBeforeTableRecs = System.Convert.ToInt32(dstCommandRowCount.ExecuteScalar());
        Console.WriteLine("Destination Table {0} Before Insert, Records = {1}", dstTable, dstBeforeTableRecs);
      }
      
      // Get data from the source as a SqlDataReader.
      SqlCommand commandSourceData = new SqlCommand
      (
        srcQuery
      , sourceConnection
      );
      commandSourceData.CommandTimeout = 0;
      SqlDataReader reader = commandSourceData.ExecuteReader();

      // Open the destination connection
      using (SqlConnection destinationConnection = new SqlConnection(dstConnStr))
      {
        destinationConnection.Open();
        
        // Set up the bulk copy object.  
        // Note that the column positions in the source 
        // data reader match the column positions in  
        // the destination table so there is no need to 
        // map columns. 
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
        {
          bulkCopy.DestinationTableName = dstTable;
          bulkCopy.BulkCopyTimeout = 0; // no timeout
          bulkCopy.BatchSize = batchSize; // 0 is default and commits entire result set

          try
          {
            // Write from the source to the destination.
            bulkCopy.WriteToServer(reader);
          }
          catch (Exception ex)
          {
            Console.WriteLine(ex.Message);
          }
          finally
          {
            // Close the SqlDataReader. The SqlBulkCopy 
            // object is automatically closed at the end of the using block.
            reader.Close();
          }
        }
        
        // Perform a final count on the destination table to see how many rows were added.
        
        SqlCommand commandRowCount = new SqlCommand();
        commandRowCount.Connection = destinationConnection;
        commandRowCount.CommandText = 
        "select " +
        "  count(*) " +
        "from " +
        dstTable + "  with (nolock) "
        ;
        long dstAfterTableRecs = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
        Console.WriteLine("Destination Table {0} After Insert, Records = {1}", dstTable, dstAfterTableRecs);
        Console.WriteLine("{0} rows were added.", dstAfterTableRecs - dstBeforeTableRecs);
        //Console.WriteLine("Press Enter to finish.");
        //Console.ReadLine();
      }
    }  
  
  } // BulkLoadQuery2Table()
  
  public static void SqlExtractData(String outputfile, String delim, String columnHeader, Boolean fieldsInQuotes, String queryFile, String server, String database, String user , String pass)
  {
    String dbConnectStr;
    StreamWriter pw = null;
    SqlConnection dbConn = null;
    String query = FileToString(queryFile);
    
    if (user == null)
    {
      dbConnectStr = "Server=" + server + ";Database=" + database + ";Trusted_Connection=True";
    }
    else
    {
      dbConnectStr = "Server=" + server + ";Database=" + database + ";user id=" + user + ";password=" + pass;
    }
       
    try
    {
      dbConn = new SqlConnection(dbConnectStr);
      dbConn.Open();
      Console.WriteLine("Opened database connection");
      int recs = 0;
      SqlCommand cmdSQL = new SqlCommand(query, dbConn);
      cmdSQL.CommandTimeout = 0;
      SqlDataReader dataReader = cmdSQL.ExecuteReader();
      int fieldCount = dataReader.FieldCount;
 
      String rec = null;
      String fieldSep = null;
      String col = null;
      int displayCount = 100000;

      Console.WriteLine("Execute SQL = " + query);
      Console.WriteLine("Number of columns in select stmt = " + fieldCount);

      pw = new StreamWriter(outputfile);

      if (columnHeader.Equals("Y"))
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
          {
            if (fieldsInQuotes)
              col = "\"" + col + "\"";
            rec += col;
          }
          else
          {
            rec += "";
          }
          fieldSep = delim;
        }

        recs++;

        if (recs % displayCount == 0)
          Console.WriteLine("Total records written so far = " + recs);
        pw.WriteLine(rec);
      }

      Console.WriteLine("\nTotal records written out to file = " + recs);

    }
    catch(Exception e)
    {
      Console.WriteLine(e.ToString());
    }
    finally
    {
      pw.Close();
      dbConn.Close();
    }

  } // SqlExtractData()
  
  public static long OdbcExtractData(String dbConnStr, String dbSqlStmt, String delim, String outputFile, Boolean columnHeader, Boolean fieldsInQuotes)
  {
  
    Console.WriteLine("\nMake connection to database URL = " + dbConnStr);
  
    // Connect to database
    OdbcConnection dbConn = new OdbcConnection(dbConnStr);
    try
    {
      dbConn.Open();
    }
    catch (Exception ex)
    {
      Console.WriteLine("Failed to Connect to database:" + ex.Message);
      return -1;
    }

    long recs = 0;
    StreamWriter pw = null;
    OdbcCommand cmdSQL = new OdbcCommand(dbSqlStmt, dbConn);
    cmdSQL.CommandTimeout = 0;
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
          {
            if (fieldsInQuotes)
              col = "\"" + col + "\"";
            rec += col;
          }
          else
          {
            rec += "";
          }
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
      Console.WriteLine("Disconnect from database");
      dbConn.Close();
      dbConn.Dispose();
    }
    
    return(recs);

  } // OdbcExtractData()

  public static long OdbcExtractDataRemoveControlChars(String dbConnStr, String dbSqlStmt, String delim, String outputFile, Boolean columnHeader)
  {
  
    Console.WriteLine("\nMake connection to database URL = " + dbConnStr);
  
    // Connect to database
    OdbcConnection dbConn = new OdbcConnection(dbConnStr);
    try
    {
      dbConn.Open();
    }
    catch (Exception ex)
    {
      Console.WriteLine("Failed to Connect to database:" + ex.Message);
      return -1;
    }

    long recs = 0;
    StreamWriter pw = null;
    OdbcCommand cmdSQL = new OdbcCommand(dbSqlStmt, dbConn);
    cmdSQL.CommandTimeout = 0;
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
          //String result = Regex.Replace(col, @"\r\n?|\n", " ");
          //String result = StripNonPrintableCharsFromString(col, " ");
          String result = CleanString(col);
          
          if (col != null)
            rec += result;
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
      Console.WriteLine("Disconnect from database");
      dbConn.Close();
      dbConn.Dispose();
    }
    
    return(recs);

  } // OdbcExtractDataRemoveControlChars()

  public static long OdbcExtractDataReplaceNewLines(String dbConnStr, String dbSqlStmt, String delim, String outputFile, Boolean columnHeader, String replaceChars)
  {
  
    Console.WriteLine("\nMake connection to database URL = " + dbConnStr);
  
    // Connect to database
    OdbcConnection dbConn = new OdbcConnection(dbConnStr);
    try
    {
      dbConn.Open();
    }
    catch (Exception ex)
    {
      Console.WriteLine("Failed to Connect to database:" + ex.Message);
      return -1;
    }

    long recs = 0;
    StreamWriter pw = null;
    OdbcCommand cmdSQL = new OdbcCommand(dbSqlStmt, dbConn);
    cmdSQL.CommandTimeout = 0;
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
          String result = Regex.Replace(col, @"\r\n?|\n", replaceChars);
          //String result = StripNonPrintableCharsFromString(col, " ");
          //String result = CleanString(col);
          
          if (col != null)
            rec += result;
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
      Console.WriteLine("Disconnect from database");
      dbConn.Close();
      dbConn.Dispose();
    }
    
    return(recs);

  } // OdbcExtractDataReplaceNewLines()
  
  /* 
   * Excel2Csv
   *
   * Generates a CSV file from an Excel spreadsheet and includes column header.
   */
  public static void Excel2Csv(String csvFile, String includeHeaderFlg, String delim, String excelFile, String worksheet, String worksheetHasHeaderFlg)
  {
    String excelConnect = null;
    String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
    String tmpDelim = "";

    if (ext.Equals("xls"))
    {
      if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
      {
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=" + excelFile + ";Extended Properties=" +
        "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      }
      else
      {
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=" + excelFile + ";Extended Properties=" +
        "\"Excel 8.0;HDR=NO;IMEX=1;\"";
      }
    }
    else if (ext.Equals("xlsx"))
    {
      if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
      {
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=" + excelFile + ";Extended Properties=" +
        "\"Excel 12.0;HDR=YES;IMEX=1;\"";
      }
      else
      {
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=" + excelFile + ";Extended Properties=" +
        "\"Excel 12.0;HDR=NO;IMEX=1;\"";
      }
    }

    String tableName = "[" + worksheet + "$]";

    StreamWriter p = new StreamWriter(csvFile);
    OleDbConnection con = new OleDbConnection(excelConnect);

    con.Open();
    Console.WriteLine("Made the connection to the spreadsheet {0}", excelFile);

    OleDbCommand command = con.CreateCommand();
    command.CommandText = "select * from " + tableName;

    OleDbDataAdapter adapter = new OleDbDataAdapter();
    adapter.SelectCommand = command;

    DataSet ds = new DataSet();
    adapter.Fill(ds, "X");
    DataTable item = ds.Tables[0];

    ArrayList v = new ArrayList();

    Console.WriteLine("The columns are:");

    foreach (DataColumn col in item.Columns)
    {
      v.Add(col.ColumnName);
      Console.WriteLine(col.ColumnName + ":" + col.DataType);

      if (includeHeaderFlg.ToUpper().Equals("T"))
      {
        p.Write(tmpDelim);
        p.Write(col.ColumnName);
        tmpDelim = delim;
      }
    }
    if (includeHeaderFlg.ToUpper().Equals("T") && worksheetHasHeaderFlg.ToUpper().Equals("T"))
      p.WriteLine();
    Console.WriteLine("The number of columns = {0}", v.Count);

    Console.WriteLine("CSV file = {0} being written out ...", csvFile);

    OleDbDataReader reader = command.ExecuteReader();
    while (reader.Read())
    {
      String s = "";
      for (int i = 0; i < v.Count; i++)
      {
        if (i == v.Count - 1) // last column?
          s += reader[(String)v[i]];
        else
          s += reader[(String)v[i]] + delim;
      }
      p.WriteLine(s);
    }

    reader.Close();
    con.Close();
    p.Close();

  } // Excel2Csv()
  
    /*
     * Excel2CsvCells
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void Excel2CsvCells(String csvFile, String includeHeaderFlg, String delim, String excelFile, String worksheet, String worksheetHasHeaderFlg, String cells)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      String tmpDelim = "";

      if (ext.Equals("xls"))
      {
        if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
        }
        else
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=NO;IMEX=1;\"";
        }
      }
      else if (ext.Equals("xlsx"))
      {
        if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";
        }
        else
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=NO;IMEX=1;\"";
        }
      }

      String tableName = "[" + worksheet + "$" + cells + "]";

      StreamWriter p = new StreamWriter(csvFile);
      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("Made the connection to the spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + tableName;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList v = new ArrayList();

      Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        v.Add(col.ColumnName);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
        
        if (includeHeaderFlg.ToUpper().Equals("T"))
        {
          p.Write(tmpDelim);
          p.Write(col.ColumnName);
          tmpDelim = delim;
        }
      }
      if (includeHeaderFlg.ToUpper().Equals("T") && worksheetHasHeaderFlg.ToUpper().Equals("T"))
        p.WriteLine();
      Console.WriteLine("The number of columns = {0}", v.Count);

      Console.WriteLine("CSV file = {0} being written out ...", csvFile);

      OleDbDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        String s = "";
        for (int i = 0; i < v.Count; i++)
        {
          if (i == v.Count - 1) // last column?
            s += reader[(String)v[i]];
          else
            s += reader[(String)v[i]] + delim;
        }
        p.WriteLine(s);
      }

      reader.Close();
      con.Close();
      p.Close();

    } // Excel2CsvCells()

    /*
     * Excel2CsvRange
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void Excel2CsvRange(String csvFile, String includeHeaderFlg, String delim, String excelFile, String worksheet, String worksheetHasHeaderFlg, String namedRange)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      String tmpDelim = "";

      if (ext.Equals("xls"))
      {
        if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=" + excelFile + ";Extended Properties=" +
            "\"Excel 8.0;HDR=YES;IMEX=1;\"";
        }
        else
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=" + excelFile + ";Extended Properties=" +
            "\"Excel 8.0;HDR=NO;IMEX=1;\"";
        }
      }
      else if (ext.Equals("xlsx"))
      {
        if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=" + excelFile + ";Extended Properties=" +
            "\"Excel 12.0;HDR=YES;IMEX=1;\"";
        }
        else
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=" + excelFile + ";Extended Properties=" +
            "\"Excel 12.0;HDR=NO;IMEX=1;\"";
        }
      }

      String tableName = "[" + namedRange + "]";

      StreamWriter p = new StreamWriter(csvFile);
      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("Made the connection to the spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + tableName;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList v = new ArrayList();

      Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        v.Add(col.ColumnName);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
        
        if (includeHeaderFlg.ToUpper().Equals("T"))
        {
          p.Write(tmpDelim);
          p.Write(col.ColumnName);
          tmpDelim = delim;
        }
      }
      if (includeHeaderFlg.ToUpper().Equals("T") && worksheetHasHeaderFlg.ToUpper().Equals("T"))
        p.WriteLine();
      Console.WriteLine("The number of columns = {0}", v.Count);

      Console.WriteLine("CSV file = {0} being written out ...", csvFile);

      OleDbDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        String s = "";
        for (int i = 0; i < v.Count; i++)
        {
          if (i == v.Count - 1) // last column?
            s += reader[(String)v[i]];
          else
            s += reader[(String)v[i]] + delim;
        }
        p.WriteLine(s);
      }

      reader.Close();
      con.Close();
      p.Close();

    } // Excel2CsvRange()
    
    /*
     * Excel2CsvSql
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void Excel2CsvSql(String csvFile, String includeHeaderFlg, String delim, String excelFile, String worksheet, String worksheetHasHeaderFlg, String sqlStmt)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      String tmpDelim = "";

      if (ext.Equals("xls"))
      {
        if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
        }
        else
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=NO;IMEX=1;\"";
        }
      }
      else if (ext.Equals("xlsx"))
      {
        if (worksheetHasHeaderFlg.ToUpper().Equals("T"))
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";
        }
        else
        {
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=NO;IMEX=1;\"";
        }
      }

      StreamWriter p = new StreamWriter(csvFile);
      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("Made the connection to the spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = sqlStmt;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList v = new ArrayList();

      Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        v.Add(col.ColumnName);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);

        if (includeHeaderFlg.ToUpper().Equals("T"))
        {
          p.Write(tmpDelim);
          p.Write(col.ColumnName);
          tmpDelim = delim;
        }
      }
      if (includeHeaderFlg.ToUpper().Equals("T") && worksheetHasHeaderFlg.ToUpper().Equals("T"))
        p.WriteLine();
      Console.WriteLine("The number of columns = {0}", v.Count);        

      Console.WriteLine("CSV file = {0} being written out ...", csvFile);

      OleDbDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        String s = "";
        for (int i = 0; i < v.Count; i++)
        {
          if (i == v.Count - 1) // last column?
            s += reader[(String)v[i]];
          else
            s += reader[(String)v[i]] + delim;
        }
        p.WriteLine(s);
      }

      reader.Close();
      con.Close();
      p.Close();

    } // Excel2CsvSql()
    
    /*
     * File2Table()
     *
     * Loads a flat file into a Sql Table.
     */
    public static void File2Table(String csvFile, String hdr, String schema, String tableName, String sqlConnStr)
    {
      String dir = Path.GetDirectoryName(csvFile);
      String csvConnect = null;
      String fileWithoutPath = Path.GetFileName(csvFile);
      
      if ( dir == null || dir.Equals("") )
      {
        csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=" + hdr + ";FMT=Delimited(,)';"; // only can use comma as delimiter here
      }
      else
      {
        csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=" + dir + ";Extended Properties='text;HDR=" + hdr + ";FMT=Delimited(,)';"; // only can use comma as delimiter here
      }
      Console.WriteLine(csvConnect);

      OleDbConnection con = new OleDbConnection(csvConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the flat file {0}", csvFile);

      OleDbCommand command = con.CreateCommand();
      //command.CommandText = "select * from [" + csvFile + "]";
      command.CommandText = "select * from [" + fileWithoutPath + "]";

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, schema + "." + tableName);
      DataTable dt = ds.Tables[0];
      
      if (dt.Rows.Count == 0)
      {
        Console.WriteLine("Zero rows selected");
        return;
      }
    
      Console.WriteLine("{0} rows selected", dt.Rows.Count );
    
      
      using (SqlConnection sqlConn = new SqlConnection(sqlConnStr))
      {
        sqlConn.Open();
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
        {
          //foreach (DataColumn c in dt.Columns)
          //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
       
          bulkCopy.DestinationTableName = dt.TableName;
          try
          {
            bulkCopy.WriteToServer(dt);
          }
          catch (Exception ex)
          {
            Console.WriteLine(ex.Message);
          }
        }
      }
  
    } // File2Table()

    /*
     * Excel2Table()
     *
     * Loads an Excel file into a Sql Table.
     */
    public static void Excel2Table(String excelFile, String worksheet, String hdr, String schema, String tableName, String sqlConnStr)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      String selectStmt = "select ";

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=" + hdr + ";IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=" + hdr + ";IMEX=1;\"";

      String sheet = "[" + worksheet + "$]";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, schema + "." + tableName);
      DataTable dt = ds.Tables[0];
      
      if (dt.Rows.Count == 0)
      {
        Console.WriteLine("Zero rows selected");
        return;
      }
    
      Console.WriteLine("{0} rows selected", dt.Rows.Count );
      
      Console.WriteLine("Columns are:");
      foreach(DataColumn col in dt.Columns)
      {
        Console.WriteLine(col.ColumnName.ToString());
      }
      
      if (hdr.Equals("YES"))
      {
        String comma = null;
        foreach(DataColumn col in dt.Columns)
        {
          // Remove any generated headers even though HDR=YES
          if (col.ColumnName.ToString().Substring(0,2).Equals("F1") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F2") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F3") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F4") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F5") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F6") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F7") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F8") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F9") )
          {
            continue;
          }
          selectStmt = selectStmt + comma + "[" + col.ColumnName + "]";
          comma = ",";
        }
        Console.WriteLine("select stmt = {0}", selectStmt);
        command.CommandText = selectStmt + " from " + sheet;
        ds = new DataSet();
        adapter.Fill(ds, schema + "." + tableName);
        dt = ds.Tables[0];
      }
      
      using (SqlConnection sqlConn = new SqlConnection(sqlConnStr))
      {
        sqlConn.Open();
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
        {
          //foreach (DataColumn c in dt.Columns)
          //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
       
          bulkCopy.DestinationTableName = dt.TableName;
          try
          {
            bulkCopy.WriteToServer(dt);
          }
          catch (Exception ex)
          {
            Console.WriteLine(ex.Message);
          }
        }
      }
    
    } // Excel2Table()

    /*
     * Excel2TableAddAuditCols()
     *
     * Loads an Excel file into a Sql Table.
     */
    public static void Excel2TableAddAuditCols(String excelFile, String worksheet, String hdr, String schema, String tableName, String sqlConnStr)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      String selectStmt = "select ";

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=" + hdr + ";IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=" + hdr + ";IMEX=1;\"";

      String sheet = "[" + worksheet + "$]";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, schema + "." + tableName);
      DataTable dt = ds.Tables[0];
      
      if (dt.Rows.Count == 0)
      {
        Console.WriteLine("Zero rows selected");
        return;
      }
    
      Console.WriteLine("{0} rows selected", dt.Rows.Count );
      
      Console.WriteLine("Columns are:");
      foreach(DataColumn col in dt.Columns)
      {
        Console.WriteLine(col.ColumnName.ToString());
      }
      
      if (hdr.Equals("YES"))
      {
        String comma = null;
        foreach(DataColumn col in dt.Columns)
        {
          // Remove any generated headers even though HDR=YES
          if (col.ColumnName.ToString().Substring(0,2).Equals("F1") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F2") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F3") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F4") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F5") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F6") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F7") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F8") ||
              col.ColumnName.ToString().Substring(0,2).Equals("F9") )
          {
            continue;
          }
          selectStmt = selectStmt + comma + "[" + col.ColumnName + "] ";
          comma = ",";
        }
        
        String filename = Path.GetFileName(excelFile);
        //String loadDate = DateTime.Now.ToString("yyyyMMdd hh:mm:ss");
        String loadDate = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
        //selectStmt = selectStmt + ", cast('" + loadDate + "' as datetime), '" + filename + "'"; 
        selectStmt = selectStmt + ", '" + loadDate + "', '" + filename + "'"; 
        
        Console.WriteLine("select stmt = {0}", selectStmt);
        command.CommandText = selectStmt + " from " + sheet;
        ds = new DataSet();
        adapter.Fill(ds, schema + "." + tableName);
        dt = ds.Tables[0];
      }
      
      using (SqlConnection sqlConn = new SqlConnection(sqlConnStr))
      {
        sqlConn.Open();
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn))
        {
          //foreach (DataColumn c in dt.Columns)
          //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
       
          bulkCopy.DestinationTableName = dt.TableName;
          try
          {
            bulkCopy.WriteToServer(dt);
          }
          catch (Exception ex)
          {
            Console.WriteLine(ex.Message);
          }
        }
      }
    
    } // Excel2TableAddAuditCols()

    public static void ExcelFilesInDir2Table(String dir, String filePattern, String worksheet, String hdr, String schema, String tableName, String sqlConnStr)
    {
      String[] fileList = Directory.GetFiles(dir, filePattern);
      String dstFile = null;
            
      foreach(String file in fileList)
      {
        dstFile = dir + Path.DirectorySeparatorChar + Path.GetFileName(file);
        Console.WriteLine("Loading file = {0}", file);
        Excel2Table(dstFile, worksheet, hdr, schema, tableName, sqlConnStr);
      }
      
    } // ExcelFilesInDir2Table()
    
    public static void FilesInDir2Table(String dir, String filePattern, String hdr, String schema, String tableName, String sqlConnStr)
    {
      String[] fileList = Directory.GetFiles(dir, filePattern);
      String dstFile = null;
            
      foreach(String file in fileList)
      {
        dstFile = dir + Path.DirectorySeparatorChar + Path.GetFileName(file);
        Console.WriteLine("Loading file = {0}", file);
        File2Table(dstFile, hdr, schema, tableName, sqlConnStr);
      }
    
    } // FilesInDir2Table()
    
  /*
   * Adds the FILENAME and LOAD_DT columns to the header record.
   * Adds the filename and load_dt to each record for audit purposes. 
   * 
   */
  public static void AddAuditInfoToFile(String srcFilePath, String dstFilePath, String fieldDelimiter, int firstRow)
  {
    StreamWriter pw;
    String[] line;
    String loadDate = DateTime.Now.ToString("yyyyMMdd hh:mm:ss");
    String s;
    String lastPart;
    String lastPartEndsWithDelim;
    int    startLine = 0;

    if (!FileExists(srcFilePath))
    {
      Console.WriteLine("File = {0} does not exist", srcFilePath);
      return;
    }
    
    pw = new StreamWriter(dstFilePath);
    line = ReadFileIntoArray(srcFilePath);
      
    if (firstRow == 2)
    {
      // adjust header record
      startLine = 1;
      s = line[0];
      if (s != null && s != "")
      {
        if (s[s.Length - 1].ToString() == fieldDelimiter)
          pw.WriteLine(s + "Filename" +  fieldDelimiter + "LoadDate");          
        else
          pw.WriteLine(s + fieldDelimiter + "Filename" + fieldDelimiter + "LoadDate");
      }
    }
    
    lastPart = fieldDelimiter + Path.GetFileName(srcFilePath) + fieldDelimiter + loadDate;
    lastPartEndsWithDelim = Path.GetFileName(srcFilePath) + fieldDelimiter + loadDate;
    for (int i = startLine; i < line.Length; i++)
    {
      s = line[i];
      if (s[s.Length - 1].ToString() == fieldDelimiter)
        pw.WriteLine(s + lastPartEndsWithDelim);      
      else
        pw.WriteLine(s + lastPart);      
    }
        
    pw.Close();
    
  } // AddAuditInfoToFile()
  
  public static void AddAuditInfoToFilesInDir(String inputDir, String filePattern, String outputDir, String fieldDelimiter, int firstRow)
  {
    String[] filePaths = Directory.GetFiles(inputDir, filePattern); 
    String loadDate = DateTime.Now.ToString("yyyyMMdd hh:mm:ss");
    String dstFilePath;
    String srcFilePath;
           
    if (filePaths.Length == 0)
    {
      Console.WriteLine("No files in '{0}' matching filePattern = {1}", inputDir, filePattern);
      return;
    }
      
    foreach (String fp in filePaths)
    {
      srcFilePath = inputDir + "\\" + Path.GetFileName(fp);
      dstFilePath = outputDir + "\\" + Path.GetFileName(fp) + "2";
      AddAuditInfoToFile(srcFilePath, dstFilePath, fieldDelimiter, firstRow);   
    }
  
  } // AddAuditInfoToFilesInDir()
  
  /*
   * GenExcelFromAccess
   *
   * Create Excel file from a SQL Server table.
   */
  public static void GenExcelFromMSSQL(String oleDbConnectStr, String tableName, String excelFile)
  {
    String fileNameNoExt = Path.GetFileNameWithoutExtension(excelFile);
    
    OleDbConnection con = new OleDbConnection(oleDbConnectStr);
    con.Open();
    Console.WriteLine("\nMade the connection to the database");
    
    OleDbCommand cmd = con.CreateCommand();
    cmd.CommandText = "select * from " + tableName;
  
    OleDbDataAdapter adapter = new OleDbDataAdapter();
    adapter.SelectCommand = cmd;
  
    DataSet ds = new DataSet();
    adapter.Fill(ds, tableName);
    //ds.WriteXml(@"my.csv");
    DataTable dt = ds.Tables[0];
          
    Console.WriteLine("{0} rows selected", dt.Rows.Count );
        
    Console.WriteLine("Columns are:");
    foreach(DataColumn col in dt.Columns)
    {
      Console.WriteLine(col.ColumnName.ToString());
    }
    
    // Bind table data to Stream Writer to export data to respective folder
    
    StreamWriter wr = new StreamWriter(fileNameNoExt + ".csv");

    // Write Columns to file

    for (int i = 0; i < dt.Columns.Count; i++)
    {
      if (i == dt.Columns.Count - 1)
        wr.Write(dt.Columns[i].ToString().ToUpper());
      else
        wr.Write(dt.Columns[i].ToString().ToUpper() + ",");

    }
    wr.WriteLine();

    //write rows to file
    for (int i = 0; i < (dt.Rows.Count); i++)
    {
      for (int j = 0; j < dt.Columns.Count; j++)
      {
        if (dt.Rows[i][j] != null)
        {
          if (j == dt.Columns.Count - 1)
            wr.Write("\"" + Convert.ToString(dt.Rows[i][j]) + "\"");
          else
            wr.Write("\"" + Convert.ToString(dt.Rows[i][j]) + "\"" + ",");
        }
        //else
        //{
        //  wr.Write(",");
        //}
      }
      wr.WriteLine();
    }
    
    wr.Close();  
            
    string FileDelimiter = ","; 
    string CreateTableStatement = "";
    string ColumnList = "";        
    
    //Read first line (Header)
    System.IO.StreamReader file = new System.IO.StreamReader(fileNameNoExt + ".csv");
    CreateTableStatement = (" Create Table [" + fileNameNoExt + "] ([" + file.ReadLine().Replace(FileDelimiter, "] Text,[")) + "] Text)";
    file.Close();
    Console.WriteLine(CreateTableStatement);
    
    File.Delete(excelFile);
                       
    // Construct ConnectionString for Excel
    string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileNameNoExt + ";" + "Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
    OleDbConnection Excel_OLE_Con = new OleDbConnection();
    OleDbCommand Excel_OLE_Cmd = new OleDbCommand();
    
    Excel_OLE_Con.ConnectionString = connstring;
    
    Console.WriteLine("Try and Open connection");
    Excel_OLE_Con.Open();
    Excel_OLE_Cmd.Connection = Excel_OLE_Con;
    
    // Use OLE DB Connection and Create Excel Sheet
    Console.WriteLine("Execute create table statement for excel file ...");
    Excel_OLE_Cmd.CommandText = CreateTableStatement;
    Excel_OLE_Cmd.ExecuteNonQuery();
    Console.WriteLine("Completed Execute create table statement for excel file ...");
        
    //Writing Data of File to Excel Sheet in Excel File
    string line;
    int counter = 0;
    
    System.IO.StreamReader SourceFile = new System.IO.StreamReader(fileNameNoExt + ".csv");
    
    while ((line = SourceFile.ReadLine()) != null)
    {
      if (counter == 0)
      {
        ColumnList = "[" + line.Replace(FileDelimiter, "],[") + "]";
        Console.WriteLine("ColumnList = {0}", ColumnList);
      }
      else
      {
        //string query = "Insert into [" + fileNameNoExt + "] (" + ColumnList  + ") VALUES('" + line.Replace(FileDelimiter, "','") + "')";
        string query = "Insert into [" + fileNameNoExt + "] (" + ColumnList  + ") VALUES(" + line.Replace("\"", "'") + ")";
        //Console.WriteLine(query);
        var command = query;
        Excel_OLE_Cmd.CommandText = command;
        Excel_OLE_Cmd.ExecuteNonQuery();
      }
      counter++;
    }
    Excel_OLE_Con.Close();
    SourceFile.Close();
    File.Delete(fileNameNoExt + ".txt");
    
    con.Close();
  
  } // GenExcelFromMSSQL()
  

} // DataUtil

} // DataLib
