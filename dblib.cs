/*
 * dblib.cs
 *
 * Craig Nobili
 */
using System;
using System.IO;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Collections;

namespace dblib
{

  public class DbUtil
  {

    // Private Members

   private static Boolean dbg = false;

    private static String accConnect;
    private static String oraConnect;

    /*
     * AccInitialize
     *
     * Sets the Access database connection string.
     */
    public static void AccInitialize(String dbFullPath)
    {
      accConnect = "Provider=Microsoft.JET.OLEDB.4.0; data source=" + dbFullPath;

    } // AccInitialize()

    /*
     * AccInitialize2
     *
     * Sets the Access database connection string.
     */
    public static String AccInitialize2(String dbFullPath)
    {
      String ext = dbFullPath.Substring(dbFullPath.IndexOf(".") + 1);

      if (ext.Equals("mdb"))
        accConnect = "Provider=Microsoft.JET.OLEDB.4.0; data source=" + dbFullPath;
      else if (ext.Equals("accdb"))
        accConnect ="Provider=Microsoft.ACE.OLEDB.12.0; data source=" + dbFullPath + ";Persist Security Info=False";

      return(accConnect);

    } // AccInitialize2()

    /*
     * OraInitialize
     *
     * Sets the Oracle database connection string.
     */
    public static void OraInitialize(String dbName, String userid, String pwd)
    {
      oraConnect = "Data Source=" + dbName + "; User ID=" + userid + "; Password=" + pwd;

    } // OraInitialize()

    /*
     * GetSqlFromFile
     *
     * Reads a file containing a SQL statement to execute.
     */
    public static String GetSqlFromFile(String filename)
    {
      String line;
      StringBuilder sb = new StringBuilder();
      StreamReader f = new StreamReader(filename);
      String newLine = null;

      while ((line = f.ReadLine()) != null)
      {
        if (line.Length >= 2 && line.Substring(0, 2) == "--") continue; // ignore comments
        sb.Append(newLine);
        sb.Append(line);
        newLine = "\n";
      }

      f.Close();

      Console.WriteLine("Got the following SQL Statement from file <{0}>", filename);
      Console.WriteLine(sb);

      return sb.ToString();

    } // GetSqlFromFile()

    /*
     * ExcelGenCsvFile
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void ExcelGenCsvFile(String csvFile, String delim, String excelFile, String worksheet)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

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
      }
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

    } // ExcelGenCsvFile()

    /*
     * ExcelGenCsvFile
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void ExcelGenCsvFile(String csvFile, String delim, String excelFile, String worksheet, String cells)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

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
      }
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

    } // ExcelGenCsvFile()

    /*
     * ExcelRangeGenCsvFile
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void ExcelRangeGenCsvFile(String csvFile, String delim, String excelFile, String namedRange)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

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
      }
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

    } // ExcelRangeGenCsvFile()

    /*
     * ExcelSqlGenCsvFile
     *
     * Generates a CSV file from an Excel spreadsheet.
     */
    public static void ExcelSqlGenCsvFile(String csvFile, String delim, String excelFile, String sqlStmt)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

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
      }
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

    } // ExcelSqlGenCsvFile()

    /*
     * ExcelGenCsvFile2
     *
     * Generates a CSV file from an Excel spreadsheet and includes column header.
     */
    public static void ExcelGenCsvFile2(String csvFile, String delim, String excelFile, String worksheet)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      String tmpDelim = "";

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

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

        p.Write(tmpDelim);
        p.Write(col.ColumnName);
        tmpDelim = delim;
      }
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

    } // ExcelGenCsvFile2()

    /*
     * AccGenCsvFile
     *
     * Generates a CSV file from an Access Database using the SQL
     * statement passed in.
     */
    public static void AccGenCsvFile(String csvFile, String delim, String sqlStmt)
    {
      StreamWriter p = new StreamWriter(csvFile);
      OleDbConnection con = new OleDbConnection(accConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");

      OleDbCommand command = con.CreateCommand();
      command.CommandText = sqlStmt;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList v = new ArrayList();

      Console.WriteLine("The following SQL was passed in:");
      Console.WriteLine(sqlStmt);
      if (dbg)
        Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        v.Add(col.ColumnName);
        if (dbg)
          Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
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

    } // AccGenCsvFile()

    /*
     * OraGenCsvFile
     *
     * Generates a CSV file from an Oracle Database using the SQL
     * statement passed in.
     */
    public static void OraGenCsvFile(String csvFile, String delim, String sqlStmt)
    {
      StreamWriter p = new StreamWriter(csvFile);
      OracleConnection con = new OracleConnection(oraConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");

      OracleCommand command = con.CreateCommand();
      command.CommandText = sqlStmt;

      OracleDataAdapter adapter = new OracleDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList v = new ArrayList();

      Console.WriteLine("The following SQL was passed in:");
      Console.WriteLine(sqlStmt);
      Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        v.Add(col.ColumnName);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      Console.WriteLine("The number of columns = {0}", v.Count);

      Console.WriteLine("CSV file = {0} being written out ...", csvFile);

      OracleDataReader reader = command.ExecuteReader();
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

    } // OraGenCsvFile()

    /*
     * AccExecNonQuery
     *
     * Executes a non-query SQL statement against Access.
     */
    public static void AccExecNonQuery(String sqlStmt)
    {
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();

      OleDbCommand cmd = con.CreateCommand();
      cmd.CommandText = sqlStmt;

      if (dbg)
      {
        Console.WriteLine("\nMade the connection to the database");
        Console.WriteLine("\nAbout to execute the following SQL Statement:");
        Console.WriteLine(cmd.CommandText);
      }

      cmd.ExecuteNonQuery();
      con.Close();

    } // AccExecNonQuery()

    /*
     * OraExecNonQuery
     *
     * Executes a non-query SQL statement against Oracle.
     */
    public static void OraExecNonQuery(String sqlStmt)
    {
      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();

      OracleCommand cmd = con.CreateCommand();
      if (sqlStmt.Substring(sqlStmt.Length - 1, 1) == ";")
        cmd.CommandText = sqlStmt.Substring(0, sqlStmt.Length - 2);
      else
        cmd.CommandText = sqlStmt;

      if (dbg)
      {
        Console.WriteLine("\nMade the connection to the database");
        Console.WriteLine("\nAbout to execute the following SQL Statement:");
        Console.WriteLine(cmd.CommandText);
      }

      cmd.ExecuteNonQuery();
      con.Close();

    } // OraExecNonQuery()

    /*
     * OraExecScalarQueryStr
     *
     * Execute a query that returns string against Oracle.
     */
    public static String OraExecScalarQueryStr(String sqlStmt)
    {
      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();

      OracleCommand cmd = con.CreateCommand();
      if (sqlStmt.Substring(sqlStmt.Length - 1, 1) == ";")
        cmd.CommandText = sqlStmt.Substring(0, sqlStmt.Length - 2);
      else
        cmd.CommandText = sqlStmt;

      if (dbg)
      {
        Console.WriteLine("\nMade the connection to the database");
        Console.WriteLine("\nAbout to execute the following SQL Statement:");
        Console.WriteLine(cmd.CommandText);
      }

      String str = (String) cmd.ExecuteScalar();
      con.Close();

      return str;

    } // OraExecScalarQueryStr()

    /*
     * AccCreateTableFromExcel
     *
     * Creates a table in Access based on the Excel spreadsheet passed in.
     */
    public static void AccCreateTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " decimal(20,4) \n";
        else
          ddl = ddl + " varchar(255) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar(255) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      AccExecNonQuery(ddl);
      con.Close();

    } // AccCreateTableFromExcel()

    public static void AccCreateTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      AccCreateTableFromExcel(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * AccCreateTableFromExcel2
     *
     * Creates a table in Access based on the Excel spreadsheet passed in.
     * Uses DateTime data type for date columns, instead of defaulting to text.
     */
    public static void AccCreateTableFromExcel2(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " decimal(20,4) \n";
        else if (colType[i].ToString().Equals("System.DateTime"))
          ddl = ddl + " datetime \n";
        else
          ddl = ddl + " varchar(255) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar(255) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      AccExecNonQuery(ddl);
      con.Close();

    } // AccCreateTableFromExcel2()

    public static void AccCreateTableFromExcel2(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      AccCreateTableFromExcel2(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * OraCreateTableFromExcel
     *
     * Creates a table in Oracle based on the Excel spreadsheet passed in.
     */
    public static void OraCreateTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " number \n";
        else
          ddl = ddl + " varchar2(256) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar(256) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      OraExecNonQuery(ddl);
      con.Close();

    } // OraCreateTableFromExcel()

    public static void OraCreateTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      OraCreateTableFromExcel(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * OraCreateTableFromExcel2
     *
     * Creates a table in Oracle based on the Excel spreadsheet passed in.
     * Uses DateTime data type for date columns, instead of defaulting to text.
     */
    public static void OraCreateTableFromExcel2(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " number \n";
        else if (colType[i].ToString().Equals("System.DateTime"))
          ddl = ddl + " date \n";
        else
          ddl = ddl + " varchar2(256) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar2(255) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      OraExecNonQuery(ddl);
      con.Close();

    } // OraCreateTableFromExcel2()

    public static void OraCreateTableFromExcel2(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      OraCreateTableFromExcel2(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * OraGenExtTabFromExcel
     *
     * Generates the SQL for an external table in Oracle based on the Excel spreadsheet passed in.
     * Uses DateTime data type for date columns, instead of defaulting to text.
     */
    public static void OraGenExtTabFromExcel(String excelFile, String excelWorksheet, String sqlFile, String tableName, String delim, String dirObjName, String location)
    {
      StreamWriter p = new StreamWriter(sqlFile);
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
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
          colName.Add(col.ColumnName.ToLower());
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
          ddl = ddl + "  " + colName[i].ToString().PadRight(30, ' ');
        else
          ddl = ddl + ", " + colName[i].ToString().PadRight(30, ' ');

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " number \n";
        else if (colType[i].ToString().Equals("System.DateTime"))
          ddl = ddl + " date \n";
        else
          ddl = ddl + " varchar2(4000) \n";
      }
      ddl = ddl + ")";
      p.WriteLine(ddl);

      p.WriteLine("organization external");
      p.WriteLine("(");
      p.WriteLine("  type oracle_loader");
      p.WriteLine("  default directory " + dirObjName.ToUpper());
      p.WriteLine("  access parameters");
      p.WriteLine("  (");
      p.WriteLine("    records delimited by NEWLINE");
      p.WriteLine("    skip 1");
      p.WriteLine("    nobadfile");
      p.WriteLine("    nodiscardfile");
      p.WriteLine("    nologfile");
      p.WriteLine("    fields terminated by '" + delim + "'");
      p.WriteLine("    missing field values are null");
      p.WriteLine("    reject rows with all null fields");
      p.WriteLine("  )");
      p.WriteLine("  location ('" + location + "')");
      p.WriteLine(")");
      p.WriteLine("reject limit unlimited");
      p.WriteLine("parallel");
      p.WriteLine(";");

      //OraExecNonQuery(ddl);
      con.Close();

      p.Close();

    } // OraGenExtTabFromExcel()

    /*
     * OraGenExtTabFromCsv
     *
     * Generates the SQL for an external table in Oracle based on the flat fle passed in.
     * Uses DateTime data type for date columns, instead of defaulting to text.
     */
    public static void OraGenExtTabFromCsv(String csvFile, String sqlFile, String tableName, String delim, String dirObjName)
    {
      StreamWriter p = new StreamWriter(sqlFile);

      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
      Console.WriteLine(csvConnect);

      OleDbConnection con = new OleDbConnection(csvConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from [" + csvFile + "]";

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");

      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      Console.WriteLine("\nThe columns are:");

      foreach (DataColumn col in item.Columns)
      {

        colName.Add(col.ColumnName.ToLower());
        colType.Add(col.DataType);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate create table ...\n");

      String ddl = "Create table " + tableName + "\n" + "( \n";

      for (int i =0; i < colName.Count; i++)
      {
        if (i == 0)
          ddl = ddl + "  " + colName[i].ToString().PadRight(30, ' ');
        else
          ddl = ddl + ", " + colName[i].ToString().PadRight(30, ' ');

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " number \n";
        else if (colType[i].ToString().Equals("System.DateTime"))
          ddl = ddl + " date \n";
        else
          ddl = ddl + " varchar2(4000) \n";
      }
      ddl = ddl + ")";
      p.WriteLine(ddl);

      p.WriteLine("organization external");
      p.WriteLine("(");
      p.WriteLine("  type oracle_loader");
      p.WriteLine("  default directory " + dirObjName.ToUpper());
      p.WriteLine("  access parameters");
      p.WriteLine("  (");
      p.WriteLine("    records delimited by NEWLINE");
      p.WriteLine("    skip 1");
      p.WriteLine("    nobadfile");
      p.WriteLine("    nodiscardfile");
      p.WriteLine("    nologfile");
      p.WriteLine("    fields terminated by '" + delim + "'");
      p.WriteLine("    missing field values are null");
      p.WriteLine("    reject rows with all null fields");
      p.WriteLine("  )");
      p.WriteLine("  location ('" + csvFile + "')");
      p.WriteLine(")");
      p.WriteLine("reject limit unlimited");
      p.WriteLine("parallel");
      p.WriteLine(";");

      //OraExecNonQuery(ddl);
      con.Close();

      p.Close();

    } // OraGenExtTabFromCsv()

    /*
     * AccCreateTableFromCsv
     *
     * Creates a table in Access based on the csv file passed in.
     */
    public static void AccCreateTableFromCsv(String csvFile, String tableName, Boolean sourceCol)
    {
      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
      Console.WriteLine(csvConnect);

      OleDbConnection con = new OleDbConnection(csvConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from [" + csvFile + "]";

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      Console.WriteLine("\nThe columns are:");

      foreach (DataColumn col in item.Columns)
      {
        colName.Add(col.ColumnName);
        colType.Add(col.DataType);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " decimal(20,4) \n";
        else
          ddl = ddl + " varchar(255) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar(255) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      AccExecNonQuery(ddl);
      con.Close();

    } // AccCreateTableFromCsv()

    /*
     * OraCreateTableFromCsv
     *
     * Creates a table in Oracle based on the csv file passed in.
     */
    public static void OraCreateTableFromCsv(String csvFile, String tableName, Boolean sourceCol)
    {
      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
      Console.WriteLine(csvConnect);

      OleDbConnection con = new OleDbConnection(csvConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from [" + csvFile + "]";

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      Console.WriteLine("\nThe columns are:");

      foreach (DataColumn col in item.Columns)
      {
        colName.Add(col.ColumnName);
        colType.Add(col.DataType);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " number \n";
        else
          ddl = ddl + " varchar2(256) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar2(255) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);

      OraExecNonQuery(ddl);
      con.Close();

    } // OraCreateTableFromCsv()

    /*
     * OraGenCrTabFromCsv
     *
     * Creates table DDL for Oracle based on the csv file passed in.
     */
    public static void OraGenCrTabFromCsv(String csvFile, String tableName,
                                          Boolean sourceCol, String outputFile)
    {
      StreamWriter pw = new StreamWriter(outputFile);

      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
      Console.WriteLine(csvConnect);

      OleDbConnection con = new OleDbConnection(csvConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "select * from [" + csvFile + "]";

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      Console.WriteLine("\nThe columns are:");

      foreach (DataColumn col in item.Columns)
      {
        colName.Add(col.ColumnName);
        colType.Add(col.DataType);
        Console.WriteLine(col.ColumnName + ":" + col.DataType);
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

        if (colType[i].ToString().Equals("System.Double") || colType[i].ToString().Equals("Decimal"))
          ddl = ddl + " number \n";
        else
          ddl = ddl + " varchar2(256) \n";
      }
      if (sourceCol)
        ddl = ddl + ", source varchar2(255) \n";
      ddl = ddl + ")";

      Console.WriteLine(ddl);
      pw.WriteLine(ddl);

      //OraExecNonQuery(ddl);
      con.Close();
      pw.Close();

    } // OraGenCrTabFromCsv()

    /*
     * AccLoadTableFromExcel
     *
     * Loads table in Access based on the Excel spreadsheet passed in.
     */
    public static void AccLoadTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

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
      //Console.WriteLine("Skipping column {0}", col.ColumnName);
      }
      else
      {
          colName.Add(col.ColumnName);
          colType.Add(col.DataType);
          //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();

      OleDbCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
          {
            if ( reader[(String)colName[i]].ToString().Equals("") )
              s = s + "  " + "null \n";
            else
              s = s + "  '" + reader[(String)colName[i]] + "'\n";
          }
          else
          {
            if ( reader[(String)colName[i]].ToString().Equals("") )
              s = s + ", " + "null \n";
            else
              s = s + ", '" + reader[(String)colName[i]] + "'\n";
          }
        }
        if (sourceCol)
          s = s + ", '" + excelFile + "_" + excelWorksheet + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //AccExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // AccLoadTableFromExcel()

    public static void AccLoadTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      AccLoadTableFromExcel(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * OraLoadTableFromExcel
     *
     * Loads table in Oracle based on the Excel spreadsheet passed in.
     */
    public static void OraLoadTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

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
      //Console.WriteLine("Skipping column {0}", col.ColumnName);
      }
      else
      {
          colName.Add(col.ColumnName);
          colType.Add(col.DataType);
          //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();

      OracleCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          // Allow implicit conversion to number, makes it easier when value is null
          if (i == 0)
            s = s + "  '" + reader[(String)colName[i]] + "'\n";
          else
            s = s + ", '" + reader[(String)colName[i]] + "'\n";
        }
        if (sourceCol)
          s = s + ", '" + excelFile + "_" + excelWorksheet + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //OraExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // OraLoadTableFromExcel()

    public static void OraLoadTableFromExcel(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      OraLoadTableFromExcel(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * OraLoadTableFromExcel2
     *
     * Loads table in Oracle based on the Excel spreadsheet passed in.
     * Handles date columns instead of defaulting to text.
     */
    public static void OraLoadTableFromExcel2(String excelFile, String excelWorksheet, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

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
      //Console.WriteLine("Skipping column {0}", col.ColumnName);
      }
      else
      {
          colName.Add(col.ColumnName);
          colType.Add(col.DataType);
          //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();

      OracleCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (colType[i].ToString().Equals("System.DateTime"))
          {
            if (i == 0)
              s = s + "  to_date('" + reader[(String)colName[i]].ToString().Substring(0, 8) + "', 'MM/DD/YYYY') \n";
            else
              s = s + ", to_date('" + reader[(String)colName[i]].ToString().Substring(0, 8) + "', 'MM/DD/YYYY') \n";
          }
          else
          {
            // Allow implicit conversion to number, makes it easier when value is null
            if (i == 0)
              s = s + "  '" + reader[(String)colName[i]] + "'\n";
            else
              s = s + ", '" + reader[(String)colName[i]] + "'\n";
          }
        }
        if (sourceCol)
          s = s + ", '" + excelFile + "_" + excelWorksheet + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //OraExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // OraLoadTableFromExcel2()

    public static void OraLoadTableFromExcel2(String excelFile, String excelWorksheet, Boolean sourceCol)
    {
      OraLoadTableFromExcel2(excelFile, excelWorksheet, sourceCol, excelWorksheet);
    }

    /*
     * AccLoadTableFromExcel2
     *
     * Loads Access table with data from Excel.
     */
    public static void AccLoadTableFromExcel2(String tableName, Boolean sourceCol, String excelFile, String excelWorksheet)
    {
      String sqlStmt;
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      if (sourceCol)
        sqlStmt = "INSERT INTO " + tableName + " SELECT t.*, '" + excelFile + "_" + excelWorksheet + "' as source FROM [" + excelWorksheet + "$] t IN '" + excelFile + "' 'Excel 8.0;HDR=YES;IMEX=1'";
      else
        sqlStmt = "INSERT INTO " + tableName + " SELECT t.* FROM [" + excelWorksheet + "$] t IN '" + excelFile + "' 'Excel 8.0;HDR=YES;IMEX=1'";

      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // AccLoadTableFromExcel2()

    /*
     * AccLoadTableFromExcel3
     *
     * Loads Access table with data from Excel.
     */
    public static void AccLoadTableFromExcel3(String tableName, Boolean sourceCol, String excelFile, String excelWorksheet)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand excelCommand = excelCon.CreateCommand();
      excelCommand.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = excelCommand;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      String comma = null;
      String selectCols = null;
      foreach (DataColumn col in item.Columns)
      {
        // Skip generated headers even though HDR=YES
        if (
            !(col.ColumnName.ToString().Substring(0,2).Equals("F1") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F2") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F3") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F4") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F5") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F6") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F7") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F8") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F9")
           )
          )
        {
          selectCols = selectCols + comma + col.ColumnName;
          comma = ",";
        }
      }
      //Console.WriteLine("cols = <{0}>", selectCols);
      excelCon.Close();

      String sqlStmt;
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      if (sourceCol)
        sqlStmt = "INSERT INTO " + tableName + " SELECT " + selectCols + ",'" + excelFile +  "_" + excelWorksheet + "' as source FROM [" + excelWorksheet + "$] IN '" + excelFile + "' 'Excel 8.0;HDR=YES;IMEX=1'";
      else
        sqlStmt = "INSERT INTO " + tableName + " SELECT " + selectCols + " FROM [" + excelWorksheet + "$] IN '" + excelFile + "' 'Excel 8.0;HDR=YES;IMEX=1'";

      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // AccLoadTableFromExcel3()

    /*
     * AccLoadTableFromExcel4
     *
     * Loads Access table with data from Excel.
     */
    public static void AccLoadTableFromExcel4(String tableName, Boolean sourceCol, String excelFile, String excelWorksheet)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String sheet = "[" + excelWorksheet + "$]";
      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand excelCommand = excelCon.CreateCommand();
      excelCommand.CommandText = "select * from " + sheet;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = excelCommand;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      String comma = null;
      String selectCols = null;
      foreach (DataColumn col in item.Columns)
      {
        // Skip generated headers even though HDR=YES
        if (
            !(col.ColumnName.ToString().Substring(0,2).Equals("F1") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F2") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F3") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F4") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F5") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F6") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F7") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F8") ||
             col.ColumnName.ToString().Substring(0,2).Equals("F9")
           )
          )
        {
          // put brackets around column it contains spaces
          selectCols = selectCols + comma + "[" + col.ColumnName + "]";
          comma = ",";
        }
      }
      //Console.WriteLine("cols = <{0}>", selectCols);
      excelCon.Close();

      String sqlStmt;
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      if (sourceCol)
        sqlStmt = "INSERT INTO " + tableName + " SELECT " + selectCols + ",'" + excelFile +  "_" + excelWorksheet + "' as source FROM [" + excelWorksheet + "$] IN '" + excelFile + "' 'Excel 8.0;HDR=YES;IMEX=1'";
      else
        sqlStmt = "INSERT INTO " + tableName + " SELECT " + selectCols + " FROM [" + excelWorksheet + "$] IN '" + excelFile + "' 'Excel 8.0;HDR=YES;IMEX=1'";

      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // AccLoadTableFromExcel4()

    /*
     * AccLoadTableFromCsv
     *
     * Loads Access table with data from Csv file.
     */
    public static void AccLoadTableFromCsv(String accessDb, String tableName, String csvFile)
    {
      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";

      OleDbConnection con = new OleDbConnection(csvConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "INSERT INTO " + tableName + " IN '" + accessDb + "' select * from " + csvFile;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(command.CommandText);

      command.ExecuteNonQuery();
      con.Close();

    } // AccLoadTableFromCsv()

    /*
     * AccCtasTableFromExcel
     *
     * Create table as select in Access using Excel data.
     */
    public static void AccCtasFromExcel(String tableName, String excelFile, String excelWorksheet)
    {
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      String sqlStmt = "SELECT * INTO " + tableName + " FROM [" + excelWorksheet + "$] IN '" + excelFile + "' 'Excel 8.0;'";
      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // AccCtasTableFromExcel()

    /*
     * AccAppendTableFromExcel
     *
     * Appends to table in Access based on the Excel spreadsheet passed in and SQL query..
     */
    public static void AccAppendTableFromExcel(String excelFile, String sqlStmt, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = sqlStmt;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

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
      //Console.WriteLine("Skipping column {0}", col.ColumnName);
      }
      else
      {
          colName.Add(col.ColumnName);
          colType.Add(col.DataType);
          //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();

      OleDbCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
          {
            if ( reader[(String)colName[i]].ToString().Equals("") )
              s = s + "  " + "null \n";
            else
              s = s + "  '" + reader[(String)colName[i]] + "'\n";
          }
          else
          {
            if ( reader[(String)colName[i]].ToString().Equals("") )
              s = s + ", " + "null \n";
            else
              s = s + ", '" + reader[(String)colName[i]] + "'\n";
          }
        }
        if (sourceCol)
          s = s + ", '" + excelFile + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //AccExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // AccAppendTableFromExcel()

    /*
     * OraAppendTableFromExcel
     *
     * Appends to table in Oracle based on the Excel spreadsheet passed in and SQL query.
     */
    public static void OraAppendTableFromExcel(String excelFile, String sqlStmt, Boolean sourceCol, String tableName)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = sqlStmt;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

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
      //Console.WriteLine("Skipping column {0}", col.ColumnName);
      }
      else
      {
          colName.Add(col.ColumnName);
          colType.Add(col.DataType);
          //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();

      OracleCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          // Allow implicit conversion to number, makes it easier when value is null
          if (i == 0)
            s = s + "  '" + reader[(String)colName[i]] + "'\n";
          else
            s = s + ", '" + reader[(String)colName[i]] + "'\n";
        }
        if (sourceCol)
          s = s + ", '" + excelFile + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //OraExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // OraAppendTableFromExcel()

    /*
     * AccLoadTableFromCsv
     *
     * Loads table in Access based on the CSV file passed in.
     */
    public static void AccLoadTableFromCsv(String csvFile, String tableName, Boolean sourceCol)
    {
      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
      Console.WriteLine(csvConnect);

      OleDbConnection excelCon = new OleDbConnection(csvConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + csvFile;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        colName.Add(col.ColumnName);
        colType.Add(col.DataType);
        //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();

      OleDbCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
          {
            if ( reader[(String)colName[i]].ToString().Equals("") )
              s = s + "  " + "null \n";
            else
              s = s + "  '" + reader[(String)colName[i]] + "'\n";
          }
          else
          {
            if ( reader[(String)colName[i]].ToString().Equals("") )
              s = s + ", " + "null \n";
            else
              s = s + ", '" + reader[(String)colName[i]] + "'\n";
          }
        }
        if (sourceCol)
          s = s + ", '" + csvFile + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //AccExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // AccLoadTableFromCsv()

    /*
     * OraLoadTableFromCsv
     *
     * Loads table in Oracle based on the CSV file passed in.
     */
    public static void OraLoadTableFromCsv(String csvFile, String tableName, Boolean sourceCol)
    {
      String csvConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
        "Data Source=.;Extended Properties='text;HDR=YES;FMT=Delimited(,)';";
      Console.WriteLine(csvConnect);

      OleDbConnection excelCon = new OleDbConnection(csvConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the CSV file {0}", csvFile);

      OleDbCommand command = excelCon.CreateCommand();
      command.CommandText = "select * from " + csvFile;

      OleDbDataAdapter adapter = new OleDbDataAdapter();
      adapter.SelectCommand = command;

      DataSet ds = new DataSet();
      adapter.Fill(ds, "X");
      DataTable item = ds.Tables[0];

      ArrayList colName = new ArrayList();
      ArrayList colType = new ArrayList();

      //Console.WriteLine("The columns are:");

      foreach (DataColumn col in item.Columns)
      {
        colName.Add(col.ColumnName);
        colType.Add(col.DataType);
        //Console.WriteLine(col.ColumnName + ":" + col.DataType);
      }
      //Console.WriteLine("The number of columns = {0}", colName.Count);

      Console.WriteLine("\nGenerate insert statements ...");

      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();

      OracleCommand cmd = con.CreateCommand();

      OleDbDataReader reader = command.ExecuteReader();
      int recs = 0;
      while (reader.Read())
      {
        String s = "insert into " + tableName + "\n" + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          if (i == 0)
            s = s + "  " + colName[i] + "\n";
          else
            s = s + ", " + colName[i] + "\n";
        }
        if (sourceCol)
          s = s + ", source \n";
        s = s + ")" + "\n";
        s = s + "values" + "\n";
        s = s + "(" + "\n";
        for (int i = 0; i < colName.Count; i++)
        {
          // Allow implicit conversion to number, makes it easier when value is null
          if (i == 0)
            s = s + "  '" + reader[(String)colName[i]] + "'\n";
          else
            s = s + ", '" + reader[(String)colName[i]] + "'\n";
        }
        if (sourceCol)
          s = s + ", '" + csvFile + "'\n";
        s = s + ")";
        //Console.WriteLine(s);
        Console.WriteLine("  inserting record {0}", ++recs);
        //OraExecNonQuery(s);
        cmd.CommandText = s;
        cmd.ExecuteNonQuery();
      }
      Console.WriteLine("\nTotal records inserted = {0}", recs);

      excelCon.Close();
      con.Close();

    } // OraLoadTableFromCsv()

    /*
     * GenExcelFromAccess
     *
     * Create Excel worksheet from an Access table.
     */
    public static void GenExcelFromAccess(String excelFile, String excelWorksheet, String accessTable)
    {
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      String sqlStmt = "SELECT * INTO " + "[" + excelWorksheet + "] IN '" + excelFile + "' 'Excel 8.0;'" + " FROM " + accessTable;
      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // GenExcelFromAccess()

    /*
     * GenExcelFromAccess
     *
     * Create Excel worksheet from a query against an Access table.  The sqlStmt
     * contains the generated Excel spreadsheet and worksheet inside it.
     */
    public static void GenExcelFromAccess(String sqlStmt)
    {
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // GenExcelFromAccess()

    /*
     * AppendExcelFromAccess
     *
     * Append to an Excel spreadsheet with data from an Access table.
     */
    public static void AppendExcelFromAccess(String excelFile, String excelWorksheet, String accessTable)
    {
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      Console.WriteLine("\nMade the connection to the database");

      OleDbCommand cmd = con.CreateCommand();

      String sqlStmt = "INSERT INTO [" + excelWorksheet + "$] IN '" + excelFile + "' 'Excel 8.0;'" + " SELECT * FROM " + accessTable;
      cmd.CommandText = sqlStmt;

      Console.WriteLine("\nAbout to execute the following SQL Statement:");
      Console.WriteLine(cmd.CommandText);

      cmd.ExecuteNonQuery();
      con.Close();

    } // AppendExcelFromAccess()

    public static String[] GetExcelSheetNames(string excelFile)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      OleDbConnection excelCon = new OleDbConnection(excelConnect);

      excelCon.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = excelCon.CreateCommand();

      // Get the data table containg the schema guid.
      DataTable dt = excelCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

      if(dt == null) return null;

      String[] excelSheets = new String[dt.Rows.Count];
      int i = 0;

      // Add the sheet name to the string array.
      foreach(DataRow row in dt.Rows)
      {
        excelSheets[i++] = row["TABLE_NAME"].ToString();
      }

      // Loop through all of the sheets
      for(int j = 0; j < excelSheets.Length; j++)
      {
        Console.WriteLine("Sheet = <{0}>", excelSheets[j]);
      }

      dt.Dispose();
      excelCon.Close();

      return excelSheets;

    } // GetExcelSheetNames()

    public static Boolean WorksheetExists(String excelFile, String worksheet)
    {
      String excelWorksheet = worksheet + "$"; // Excel puts a $ after the worksheet name
      // Quotes put around worksheet names containing spaces
      String excelWorksheetSpaces = "'" + excelWorksheet + "'";
      String [] sheets = GetExcelSheetNames(excelFile);

      for (int i = 0; i < sheets.Length; i++)
      {
        if (sheets[i].Equals(excelWorksheet) || sheets[i].Equals(excelWorksheetSpaces))
          return true;
      }
      return false;

    } // WorksheetExists()

    public static Boolean FileExists(string filename)
    {
      if (File.Exists(filename)) return true;
      return false;

    } // FileExists()

    // **** Less General Methods ********************

    public static void AccDropTable(String tableName)
    {
      OleDbConnection con = new OleDbConnection(accConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");

      OleDbCommand cmd = con.CreateCommand();
      cmd.CommandText = "drop table " + tableName;

      try
      {
        cmd.ExecuteNonQuery();
        Console.WriteLine("Dropped table {0}", tableName);
      }
      catch (OleDbException e)
      {
        if (e != null)
          Console.WriteLine("Table {0} does not exist", tableName);
      }
      finally
      {
        con.Close();
      }

    } // AccDropTable()

    public static void OraDropTable(String tableName)
    {
      OracleConnection con = new OracleConnection(oraConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");

      OracleCommand cmd = con.CreateCommand();
      cmd.CommandText = "drop table " + tableName;

      try
      {
        cmd.ExecuteNonQuery();
        Console.WriteLine("Dropped table {0}", tableName);
      }
      catch (OracleException e)
      {
        if (e != null)
          Console.WriteLine("Table {0} does not exist", tableName);
      }
      finally
      {
        con.Close();
      }

    } // OraDropTable()

    public static void AccCreateTable(String ddlFile)
    {
      OleDbConnection con = new OleDbConnection(accConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");
      Console.WriteLine("Get DDL from file {0}", ddlFile);

      OleDbCommand cmd = con.CreateCommand();
      cmd.CommandText = GetSqlFromFile(ddlFile);

      try
      {
        Console.WriteLine("Creating table ...");
        cmd.ExecuteNonQuery();
      }
      catch (OleDbException e)
      {
        Console.WriteLine(e);
      }
      finally
      {
        con.Close();
      }

    } // AccCreateTable()

    public static void OraCreateTable(String ddlFile)
    {
      String sqlStmt;
      OracleConnection con = new OracleConnection(oraConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");
      Console.WriteLine("Get DDL from file {0}", ddlFile);

      OracleCommand cmd = con.CreateCommand();

      sqlStmt = GetSqlFromFile(ddlFile);

      if (sqlStmt.Substring(sqlStmt.Length - 1, 1) == ";")
        cmd.CommandText = sqlStmt.Substring(0, sqlStmt.Length - 2);
      else
        cmd.CommandText = sqlStmt;

      try
      {
        Console.WriteLine("Creating table ...");
        cmd.ExecuteNonQuery();
      }
      catch (OracleException e)
      {
        Console.WriteLine(e);
      }
      finally
      {
        con.Close();
      }

    } // OraCreateTable()

    public static void AccDelTable(String tableName)
    {
      OleDbConnection con = new OleDbConnection(accConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");

      OleDbCommand cmd = con.CreateCommand();
      cmd.CommandText = "delete from " + tableName;

      try
      {
        Console.WriteLine("deleting from {0} ...", tableName);
        cmd.ExecuteNonQuery();
      }
      catch (OleDbException e)
      {
        Console.WriteLine(e);
      }
      finally
      {
        con.Close();
      }

    } // AccDelTable()

    public static void OraDelTable(String tableName)
    {
      OracleConnection con = new OracleConnection(oraConnect);

      con.Open();
      Console.WriteLine("Made the connection to the database");

      OracleCommand cmd = con.CreateCommand();
      cmd.CommandText = "delete from " + tableName;

      try
      {
        Console.WriteLine("deleting from {0} ...", tableName);
        cmd.ExecuteNonQuery();
      }
      catch (OracleException e)
      {
        Console.WriteLine(e);
      }
      finally
      {
        con.Close();
      }

    } // OraDelTable()

    // **** End Less General Methods ********************

    // ************************************** examples **************************

    public static void UpdateExcel(String excelFile, String excelWorksheet)
    {
      String excelConnect = null;
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);

      if (ext.Equals("xls"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 8.0;HDR=YES;IMEX=1;\"";
      else if (ext.Equals("xlsx"))
        excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
          "Data Source=" + excelFile + ";Extended Properties=" +
          "\"Excel 12.0;HDR=YES;IMEX=1;\"";

      String tableName = "[" + excelWorksheet + "$]";

      OleDbConnection con = new OleDbConnection(excelConnect);

      con.Open();
      Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

      OleDbCommand command = con.CreateCommand();
      command.CommandText = "update  " + tableName + " set City = 'testupdate' where ID = 1";
      command.ExecuteNonQuery();
      command.CommandText = "insert into " + tableName + " (id, city, state, upd_dt) values (6, 'testcity', 'teststate', '8/15/2008') ";
      command.ExecuteNonQuery();
      // Can't delete from Excel spreadsheet?
      //command.CommandText = "delete from " + tableName + " where ID = 2";
      //command.ExecuteNonQuery();

      con.Close();

    } // UpdateExcel()

    public static void AccInsSampleFromExcel()
    {
      int recs = 0;
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      OleDbCommand cmd = con.CreateCommand();

      // Put Excel spreadsheet filename here
      string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;
        Data Source=Book1.xls;Extended Properties=
        ""Excel 8.0;HDR=YES;""";

      string insSql;

      DbProviderFactory factory =  DbProviderFactories.GetFactory("System.Data.OleDb");

      using (DbConnection connection = factory.CreateConnection())
      {
        connection.ConnectionString = connectionString;

        using (DbCommand command = connection.CreateCommand())
        {
          // Cities$ comes from the name of the worksheet
          command.CommandText = "SELECT ID,City,State FROM [Cities$]";

          connection.Open();

          using (DbDataReader dr = command.ExecuteReader())
          {
            while (dr.Read())
            {
              // Spreadsheet column headers
              //Console.WriteLine(dr["ID"].ToString() + ":" + dr["City"].ToString() + ":" + dr["State"].ToString());
              insSql = "insert into sample1 " +
                       "( " +
                       "  ID " +
                       ", City " +
                       ", State " +
                       ") " +
                       "values " +
                       "( " +
                       "  " + dr["ID"].ToString() + " " +
                       ",  '" + dr["City"].ToString() + "' " +
                       ",  '" + dr["State"].ToString() + "' " +
                       ") "
                       ;
              //Console.WriteLine(insSql);
              cmd.CommandText = insSql;
              cmd.ExecuteNonQuery();
              Console.WriteLine("Inserted record {0}", ++recs);
            }
          }
          connection.Close();
          }
     }

     con.Close();

    } // AccInsSampleFromExcel()

    public static void AccInsSampleFromExcel2()
    {
      int recs = 0;
      OleDbConnection con = new OleDbConnection(accConnect);
      con.Open();
      OleDbCommand cmd = con.CreateCommand();

      // Put Excel spreadsheet filename here
      string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;
        Data Source=Book2.xls;Extended Properties=
        ""Excel 8.0;HDR=YES;""";

      string insSql;

      DbProviderFactory factory =  DbProviderFactories.GetFactory("System.Data.OleDb");

      using (DbConnection connection = factory.CreateConnection())
      {
        connection.ConnectionString = connectionString;

        using (DbCommand command = connection.CreateCommand())
        {
          // Cities$ comes from the name of the worksheet
          command.CommandText = "SELECT ID,City,State, upd_dt FROM [Cities$]";

          connection.Open();

          using (DbDataReader dr = command.ExecuteReader())
          {
            while (dr.Read())
            {
              // Spreadsheet column headers
              //Console.WriteLine(dr["ID"].ToString() + ":" + dr["City"].ToString() + ":" + dr["State"].ToString());
              insSql = "insert into sample " +
                       "( " +
                       "  ID " +
                       ", City " +
                       ", State " +
                       ", upd_dt " +
                       ") " +
                       "values " +
                       "( " +
                       "  " + dr["ID"].ToString() + " " +
                       ",  '" + dr["City"].ToString() + "' " +
                       ",  '" + dr["State"].ToString() + "' " +
                       ",  '" + dr["upd_dt"].ToString() + "' " +
                       ") "
                       ;
              cmd.CommandText = insSql;
              cmd.ExecuteNonQuery();
              Console.WriteLine("Inserted record {0}", ++recs);
            }
          }
          connection.Close();
          }
     }

     con.Close();

    } // AccInsSampleFromExcel2()

    public static void OraInsSampleFromExcel()
    {
      int recs = 0;
      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();
      OracleCommand cmd = con.CreateCommand();

      // Put Excel spreadsheet filename here
      string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;
        Data Source=Book1.xls;Extended Properties=
        ""Excel 8.0;HDR=YES;""";

      string insSql;

      DbProviderFactory factory =  DbProviderFactories.GetFactory("System.Data.OleDb");

      using (DbConnection connection = factory.CreateConnection())
      {
        connection.ConnectionString = connectionString;

        using (DbCommand command = connection.CreateCommand())
        {
          // Cities$ comes from the name of the worksheet
          command.CommandText = "SELECT ID,City,State FROM [Cities$]";

          connection.Open();

          using (DbDataReader dr = command.ExecuteReader())
          {
            while (dr.Read())
            {
              // Spreadsheet column headers
              //Console.WriteLine(dr["ID"].ToString() + ":" + dr["City"].ToString() + ":" + dr["State"].ToString());
              insSql = "insert into sample1 " +
                       "( " +
                       "  ID " +
                       ", City " +
                       ", State " +
                       ") " +
                       "values " +
                       "( " +
                       "  " + dr["ID"].ToString() + " " +
                       ",  '" + dr["City"].ToString() + "' " +
                       ",  '" + dr["State"].ToString() + "' " +
                       ") "
                       ;
              cmd.CommandText = insSql;
              cmd.ExecuteNonQuery();
              Console.WriteLine("Inserted record {0}", ++recs);
            }
          }
          connection.Close();
          }
     }

     con.Close();

    } // OraInsSampleFromExcel()

    public static void OraInsSampleFromExcel2()
    {
      int recs = 0;
      int dateLen;
      String dateStr;
      OracleConnection con = new OracleConnection(oraConnect);
      con.Open();
      OracleCommand cmd = con.CreateCommand();

      // Put Excel spreadsheet filename here
      string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;
        Data Source=Book2.xls;Extended Properties=
        ""Excel 8.0;HDR=YES;""";

      string insSql;

      DbProviderFactory factory =  DbProviderFactories.GetFactory("System.Data.OleDb");

      using (DbConnection connection = factory.CreateConnection())
      {
        connection.ConnectionString = connectionString;

        using (DbCommand command = connection.CreateCommand())
        {
          // Cities$ comes from the name of the worksheet
          command.CommandText = "SELECT ID,City,State, upd_dt FROM [Cities$]";

          connection.Open();

          using (DbDataReader dr = command.ExecuteReader())
          {
            while (dr.Read())
            {
              // Spreadsheet column headers
              //Console.WriteLine(dr["ID"].ToString() + ":" + dr["City"].ToString() + ":" + dr["State"].ToString());
              dateLen = dr["upd_dt"].ToString().Length;
              dateStr = dr["upd_dt"].ToString().Substring(0, dateLen - 2); // take off am/pm indicator
              insSql = "insert into sample " +
                       "( " +
                       "  ID " +
                       ", City " +
                       ", State " +
                       ", upd_dt " +
                       ") " +
                       "values " +
                       "( " +
                       "  " + dr["ID"].ToString() + " " +
                       ",  '" + dr["City"].ToString() + "' " +
                       ",  '" + dr["State"].ToString() + "' " +
                       ",  trunc(to_date('" + dateStr + "', 'MM/DD/YYYY hh24:mi:ss')) " +
                       ") "
                       ;
              //Console.WriteLine(insSql);
              cmd.CommandText = insSql;
              cmd.ExecuteNonQuery();
              Console.WriteLine("Inserted record {0}", ++recs);
            }
          }
          connection.Close();
          }
     }

     con.Close();

    } // OraInsSampleFromExcel2()

    // ************************************** end examples **************************

  } // DbUtil

} // dblib
