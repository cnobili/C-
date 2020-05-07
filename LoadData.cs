/*
 * LoadData.cs
 *
 * Table Driven data loads.
 *
 * Craig Nobili
 */

using System;
using System.Text;
using System.Configuration;
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

namespace Data
{

class DbReturnStatus
{
  public int    retCode    = 0;
  public String errMsg     = null;
  public long   rowsLoaded = 0;

} // DbReturnStatus class

class LoadData
{

  /*
   * Private Data
   */
   
  public const String PROGRAM_NAME = "LoadData";
  public const String SUCCESS = "SUCCESS";
  public const String FAILURE = "FAILURE";
  
  private static String metadataDbConnStr;
  private static String metadataDbName;
  private String loadDataName;
  private String groupName;
         
  private OdbcConnection  srcOdbcConn  = null;
  private OleDbConnection srcOleDbConn = null;
  private SqlConnection   srcSqlConn   = null;

  /*
   * Public Methods
   */
  
  public LoadData(String connStr, String dbName, String loadDataName, String groupName)
  {
    metadataDbConnStr = connStr;
    metadataDbName = dbName;
    this.loadDataName = loadDataName;
    this.groupName = groupName;

  } // LoadData()

  /*
   * ConnectToDb()
   *
   * Opens up a connection to the database.
   */
  public void ConnectToDb(String connType, String connStr)
  {
    Console.WriteLine("\nMake connection to database URL = " + connStr);

    // Connect to database
    if (connType.ToUpper().Equals("MSSQL"))
    {
      srcSqlConn = new SqlConnection(connStr);
      srcSqlConn.Open();
    }
    else if (connType.ToUpper().Equals("OLEDB"))
    {
      srcOleDbConn = new OleDbConnection(connStr);
      srcOleDbConn.Open();
    }
    else if (connType.ToUpper().Equals("ODBC"))
    {
      srcOdbcConn = new OdbcConnection(connStr);
      srcOdbcConn.Open();
    }
    else
    {
      Console.WriteLine("ERROR: unknown connection type");
    }

  } // ConnectToDb()

  /*
   *  DisconnectFromDb()
   *
   * Disconnects from the database.
   */
  public void DisconnectFromDb(String connType)
  {
    Console.WriteLine("Disconnect from database");
    if (connType.ToUpper().Equals("MSSQL"))
    {
      srcSqlConn.Close();
      srcSqlConn.Dispose();
    }
    else if (connType.ToUpper().Equals("OLEDB"))
    {
      srcOleDbConn.Close();
      srcOleDbConn.Dispose();
    }
    else if (connType.ToUpper().Equals("ODBC"))
    {
      srcOdbcConn.Close();
      srcOdbcConn.Dispose();
    }
    else
    {
      Console.WriteLine("ERROR: unknown connection type");
    }

  } // DisconnectFromDb()

  public static DbReturnStatus OleDbQuery2SqlTable(String odbcConnStr, String query, String dstConnStr, String schemaName, String dstTable, String loadType)
  {
    DbReturnStatus ret = new DbReturnStatus();
    OleDbDataAdapter da = null;
    DataTable dt = null;
    SqlBulkCopy bulkCopy = null;
    SqlConnection sqlConn = null;
    
    try
    {
      da = new OleDbDataAdapter(query, odbcConnStr);
      da.SelectCommand.CommandTimeout = 0;
      
      dt = new DataTable(schemaName + "." + dstTable);
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
        ret.retCode = 0;
        ret.errMsg = SUCCESS;
        ret.rowsLoaded = 0;
        return ret;
      }
    
      Console.WriteLine("{0} rows selected", dt.Rows.Count );
          
      if (loadType.ToUpper().Equals("FULL"))
      {
        String sqlStmt = "truncate table " + schemaName + "." + dstTable;
        SqlCommand cmd = new SqlCommand();
        SqlConnection conn = new SqlConnection(dstConnStr);
             
        conn.Open();
        cmd.Connection = conn;
        cmd.CommandText = sqlStmt;
        cmd.ExecuteNonQuery();
        conn.Close();
      }
      
      sqlConn = new SqlConnection(dstConnStr);
      sqlConn.Open();
      bulkCopy = new SqlBulkCopy(sqlConn);
        
      //foreach (DataColumn c in dt.Columns)
      //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
    
      bulkCopy.DestinationTableName = dt.TableName;
      bulkCopy.BulkCopyTimeout = 0;
      bulkCopy.WriteToServer(dt);
    }
    catch (Exception ex)
    {
      Console.WriteLine("Error = " + ex.Message);
      ret.retCode = 1;
      ret.errMsg = ex.Message;
      ret.rowsLoaded = -1;
      return ret;
    }
    finally
    {
      ret.retCode = 0;
      ret.errMsg = SUCCESS;
      ret.rowsLoaded = dt.Rows.Count;
      da.Dispose();
      bulkCopy.Close();
      sqlConn.Close();
    }
    return ret;
      
  } // OleDbQuery2SqlTable()

  public static DbReturnStatus OdbcQuery2SqlTable(String odbcConnStr, String query, String dstConnStr, String schemaName, String dstTable, String loadType)
  {
    DbReturnStatus ret = new DbReturnStatus();
    OdbcDataAdapter da = null;
    DataTable dt = null;
    SqlBulkCopy bulkCopy = null;
    SqlConnection sqlConn = null;
    
    try
    {
      da = new OdbcDataAdapter(query, odbcConnStr);
      da.SelectCommand.CommandTimeout = 0;
      
      dt = new DataTable(schemaName + "." + dstTable);
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
        ret.retCode = 0;
        ret.errMsg = SUCCESS;
        ret.rowsLoaded = 0;
        return ret;
      }
    
      Console.WriteLine("{0} rows selected", dt.Rows.Count );
          
      if (loadType.ToUpper().Equals("FULL"))
      {
        String sqlStmt = "truncate table " + schemaName + "." + dstTable;
        SqlCommand cmd = new SqlCommand();
        SqlConnection conn = new SqlConnection(dstConnStr);
             
        conn.Open();
        cmd.Connection = conn;
        cmd.CommandText = sqlStmt;
        cmd.ExecuteNonQuery();
        conn.Close();
      }
      
      sqlConn = new SqlConnection(dstConnStr);
      sqlConn.Open();
      bulkCopy = new SqlBulkCopy(sqlConn);
        
      //foreach (DataColumn c in dt.Columns)
      //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
    
      bulkCopy.DestinationTableName = dt.TableName;
      bulkCopy.BulkCopyTimeout = 0;
      bulkCopy.WriteToServer(dt);
    }
    catch (Exception ex)
    {
      Console.WriteLine("Error = " + ex.Message);
      ret.retCode = 1;
      ret.errMsg = ex.Message;
      ret.rowsLoaded = -1;
      return ret;
    }
    finally
    {
      ret.retCode = 0;
      ret.errMsg = SUCCESS;
      ret.rowsLoaded = dt.Rows.Count;
      da.Dispose();
      bulkCopy.Close();
      sqlConn.Close();
    }
    return ret;
      
  } // OdbcQuery2SqlTable()
  
  public static DbReturnStatus BulkCopyTable(String srcConnStr, String srcQuery, String dstConnStr, String dstSchema, String dstTable, int batchSize, String loadType)
  {
    DbReturnStatus ret = new DbReturnStatus();
    SqlConnection sourceConnection = null;
    SqlConnection destinationConnection = null;
    SqlDataReader reader = null;
    SqlBulkCopy bulkCopy = null;
    long dstBeforeTableRecs = -1;
    long dstAfterTableRecs = -1;
    
    try
    {
      sourceConnection = new SqlConnection(srcConnStr);
      sourceConnection.Open();
        
      destinationConnection = new SqlConnection(dstConnStr);
      destinationConnection.Open();
      
      if (loadType.ToUpper().Equals("FULL"))
      {
        String sqlStmt = "truncate table " + dstSchema + "." + dstTable;
        SqlCommand cmd = new SqlCommand();
        cmd.Connection = destinationConnection;
        cmd.CommandText = sqlStmt;
        cmd.ExecuteNonQuery();
      }
           
      // Get destination before row count
      SqlCommand commandRowCount = new SqlCommand();
      commandRowCount.CommandTimeout = 0;
      commandRowCount.Connection = destinationConnection;
      commandRowCount.CommandText = 
        "select " +
        "  count(*) " +
        "from " +
       dstSchema + "." + dstTable + "  with (nolock) "
       ;
       dstBeforeTableRecs = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
       Console.WriteLine("Destination table before row count = {0}", dstBeforeTableRecs);
       
      // Get data from the source table as a SqlDataReader.
      SqlCommand commandSourceData = new SqlCommand(srcQuery, sourceConnection);
      commandSourceData.CommandTimeout = 0;
      reader = commandSourceData.ExecuteReader();

      // Set up the bulk copy object.  
      // Note that the column positions in the query of the 
      // data reader must match the column positions in  
      // the destination table so there is no need to 
      // map columns.
      bulkCopy = new SqlBulkCopy(destinationConnection);
      bulkCopy.DestinationTableName = dstSchema + "." + dstTable;
      bulkCopy.BulkCopyTimeout = 0; // no timeout
      bulkCopy.BatchSize = batchSize; // 0 is default and commits entire result set
    
      // Write from the source to the destination.
      bulkCopy.WriteToServer(reader);
      
      // Get Destination after row count
      commandRowCount = new SqlCommand();
      commandRowCount.Connection = destinationConnection;
      commandRowCount.CommandText = 
        "select " +
        "  count(*) " +
        "from " +
           dstSchema + "." + dstTable + "  with (nolock) "
      ;
      dstAfterTableRecs = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
      Console.WriteLine("Destination table aftere row count = {0}", dstAfterTableRecs);
      Console.WriteLine("{0} rows were added.", dstAfterTableRecs - dstBeforeTableRecs);

    }
    catch (Exception ex)
    {
      Console.WriteLine("Error = " + ex.Message);
      ret.retCode = 1;
      ret.errMsg = ex.Message;
      ret.rowsLoaded = -1;
      return ret;
    }
    finally
    {
      reader.Close();
      bulkCopy.Close();
      sourceConnection.Close();
      destinationConnection.Close();
      ret.retCode = 0;
      ret.errMsg = SUCCESS;
      ret.rowsLoaded = dstAfterTableRecs - dstBeforeTableRecs;
    }
    return ret;
  
  } // BulkCopyTable()
  
    /*
     * File2Table()
     *
     * Loads a flat file into a Sql Table.
     */
    public static DbReturnStatus File2Table(String csvFile, String hdr, String schema, String tableName, String sqlConnStr, String loadType)
    {
      DbReturnStatus ret = new DbReturnStatus();
      String dir = Path.GetDirectoryName(csvFile);
      String csvConnect = null;
      String fileWithoutPath = Path.GetFileName(csvFile);
      SqlConnection sqlConn = null;
      SqlBulkCopy bulkCopy = null;
      DataTable dt = null;
      DataSet ds = null;

      try
      {
     
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
        command.CommandTimeout = 0;
        //command.CommandText = "select * from [" + csvFile + "]";
        command.CommandText = "select * from [" + fileWithoutPath + "]";

        OleDbDataAdapter adapter = new OleDbDataAdapter();
        adapter.SelectCommand = command;

        ds = new DataSet();
        adapter.Fill(ds, schema + "." + tableName);
        dt = ds.Tables[0];
      
        if (dt.Rows.Count == 0)
        {
          Console.WriteLine("Zero rows selected");
          ret.retCode = 0;
          ret.errMsg = SUCCESS;
          ret.rowsLoaded = 0;
          return ret;
        }
    
        Console.WriteLine("{0} rows selected", dt.Rows.Count );
          
        if (loadType.ToUpper().Equals("FULL"))
        {
          String sqlStmt = "truncate table " + schema + "." + tableName;
          SqlCommand cmd = new SqlCommand();
          SqlConnection conn = new SqlConnection(sqlConnStr);
             
          conn.Open();
          cmd.Connection = conn;
          cmd.CommandText = sqlStmt;
          cmd.ExecuteNonQuery();
          conn.Close();
        }
      
        sqlConn = new SqlConnection(sqlConnStr);
        sqlConn.Open();
        bulkCopy = new SqlBulkCopy(sqlConn);
        
        //foreach (DataColumn c in dt.Columns)
        //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
    
        bulkCopy.DestinationTableName = dt.TableName;
        bulkCopy.BulkCopyTimeout = 0;
        bulkCopy.WriteToServer(dt);
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error = " + ex.Message);
        ret.retCode = 1;
        ret.errMsg = ex.Message;
        ret.rowsLoaded = -1;
        return ret;
      }
      finally
      {
        ret.retCode = 0;
        ret.errMsg = SUCCESS;
        ret.rowsLoaded = dt.Rows.Count;
        ds.Dispose();
        bulkCopy.Close();
        sqlConn.Close();
      }
      return ret;

   } // File2Table()
   
   public static void GenSchemaINI(String fullPathSchemaINI, String filename, String delim, String colNameHeader)
   {
     StreamWriter sw = new StreamWriter(fullPathSchemaINI);
     sw.WriteLine("[{0}]", filename);
     sw.WriteLine("ColNameHeader={0}", colNameHeader);
     sw.WriteLine("Format=Delimited({0})", delim);
     sw.Close();
       
   } // GenSchemaINI()

   
    public static DbReturnStatus FilesInDir2Table(String dir, String filePattern, String hdr, String schema, String tableName, String sqlConnStr, String loadType, String delimiter, String connType,
      String logDbConnStr, String programName, String loadDataName, String sqlStatement, String procOrSqlBlk, String destDbName, int batchSize, String worksheet, long loadDataDefKey)
    {
      DbReturnStatus ret = null;
      String[] fileList = Directory.GetFiles(dir, filePattern);
      //String filename = null;
      long loadDataOutKey;
      
      int i = 0;
      foreach(String file in fileList)
      {
        //filename = dir + Path.DirectorySeparatorChar + Path.GetFileName(file);
        Console.WriteLine("Loading file = {0}", file);
        loadDataOutKey = LogBeg(logDbConnStr, programName, loadDataName, sqlStatement, procOrSqlBlk, DateTime.Now, sqlConnStr, destDbName, schema, tableName, loadType, batchSize, dir, filePattern, worksheet, hdr, delimiter, Path.GetFileName(file), loadDataDefKey);
        
        if (connType.ToUpper().Equals("FLATFILE"))
        {
          GenSchemaINI( dir + Path.DirectorySeparatorChar + "schema.ini",  Path.GetFileName(file), delimiter, (hdr == "Y" ? "True" : "False") );
        }
        
        if (i == 0)
        {
          if (hdr.ToUpper().Equals("N"))
          {
            Console.WriteLine("LoadDataDef has Header = N, therefore the destination table must exist in the DB");
          }
          else
          {
            if (!TableExists(schema, tableName, sqlConnStr))
            {
              CreateTableFromFile(schema, tableName, sqlConnStr, file, delimiter);
            }
          }
          ret = File2Table(file, hdr, schema, tableName, sqlConnStr, loadType);
        }
        else
        {
          // if multiple files, after the first one, the rest need to be appended
          ret = File2Table(file, hdr, schema, tableName, sqlConnStr, "APPEND");
        }
        
        LogEnd(logDbConnStr, loadDataOutKey, DateTime.Now, (ret.errMsg == SUCCESS ? SUCCESS : FAILURE), ret.rowsLoaded, (ret.errMsg == SUCCESS ? null : ret.errMsg));
        if (ret.retCode != 0) break;
        i++;
      }
      
      return(ret);
    
    } // FilesInDir2Table()
    
    /*
     * Excel2Table()
     *
     * Loads an Excel file into a Sql Table.
     */
    public static DbReturnStatus Excel2Table(String excelFile, String worksheet, String hdr, String schema, String tableName, String sqlConnStr, String loadType)
    {
      DbReturnStatus ret  = new DbReturnStatus();
      String excelConnect = null;
      
      String ext = excelFile.Substring(excelFile.LastIndexOf(".") + 1);
      
      OleDbConnection excelCon = null;
      OleDbCommand command     = null;
      OleDbDataAdapter adapter = null;
      SqlConnection sqlConn    = null;
      SqlBulkCopy bulkCopy     = null;
      DataSet ds               = null;
      DataTable dt             = null;

      try
      {
        if (ext.Equals("xls"))
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=" + excelFile + ";Extended Properties=" +
            "\"Excel 8.0;HDR=" + hdr + ";IMEX=1;\"";
        else if (ext.Equals("xlsx"))
          excelConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            "Data Source=" + excelFile + ";Extended Properties=" +
            "\"Excel 12.0;HDR=" + hdr + ";IMEX=1;\"";

        String sheet = "[" + worksheet + "$]";

        excelCon = new OleDbConnection(excelConnect);

        excelCon.Open();
        Console.WriteLine("\nMade the connection to the Spreadsheet {0}", excelFile);

        command = excelCon.CreateCommand();
        command.CommandTimeout = 0;
        command.CommandText = "select * from " + sheet;

        adapter = new OleDbDataAdapter();
        adapter.SelectCommand = command;

        ds = new DataSet();
        adapter.Fill(ds, tableName);
        dt = ds.Tables[0];

        if (dt.Rows.Count == 0)
        {
          Console.WriteLine("Zero rows selected");
          ret.retCode = 0;
          ret.errMsg = SUCCESS;
          ret.rowsLoaded = 0;
          return ret;
        }

        Console.WriteLine("{0} rows selected", dt.Rows.Count );
        
        if (loadType.ToUpper().Equals("FULL"))
        {
          String sqlStmt = "truncate table " + schema + "." + tableName;
          SqlCommand cmd = new SqlCommand();
          SqlConnection conn = new SqlConnection(sqlConnStr);
             
          conn.Open();
          cmd.Connection = conn;
          cmd.CommandText = sqlStmt;
          cmd.ExecuteNonQuery();
          conn.Close();
        }

        sqlConn = new SqlConnection(sqlConnStr);
        sqlConn.Open();
        bulkCopy = new SqlBulkCopy(sqlConn);
        
        //foreach (DataColumn c in dt.Columns)
        //  bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);

        bulkCopy.DestinationTableName = dt.TableName;
        bulkCopy.BulkCopyTimeout = 0;
        bulkCopy.WriteToServer(dt);
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error = " + ex.Message);
        ret.retCode = 1;
        ret.errMsg = ex.Message;
        ret.rowsLoaded = -1;
        return ret;
      }
      finally
      {
        ret.retCode = 0;
        ret.errMsg = SUCCESS;
        ret.rowsLoaded = dt.Rows.Count;
        ds.Dispose();
        bulkCopy.Close();
        sqlConn.Close();
      }
      return ret;
    
    } // Excel2Table()
    
    public static DbReturnStatus ExcelFilesInDir2Table(String dir, String filePattern, String hdr, String schema, String tableName, String sqlConnStr, String loadType, String delimiter, String connType,
      String logDbConnStr, String programName, String loadDataName, String sqlStatement, String procOrSqlBlk, String destDbName, int batchSize, String worksheet, long loadDataDefKey)
    {
      DbReturnStatus ret = null;
      String[] fileList = Directory.GetFiles(dir, filePattern);
      //String filename = null;
      long loadDataOutKey;
      
      int i = 0;
      foreach(String file in fileList)
      {
        //filename = dir + Path.DirectorySeparatorChar + Path.GetFileName(file);
        Console.WriteLine("Loading file = {0}", file);
        loadDataOutKey = LogBeg(logDbConnStr, programName, loadDataName, sqlStatement, procOrSqlBlk, DateTime.Now, sqlConnStr, destDbName, schema, tableName, loadType, batchSize, dir, filePattern, worksheet, hdr, delimiter, Path.GetFileName(file), loadDataDefKey);
        
        if (i == 0)
        {
          if (hdr.ToUpper().Equals("N"))
          {
            Console.WriteLine("LoadDataDef has Header = N, therefore the destination table must exist in the DB");
          }
          else
          {
            if (!TableExists(schema, tableName, sqlConnStr))
            {
              CreateTableFromExcel(schema, tableName, sqlConnStr, file, worksheet);
            }
          }
          ret = Excel2Table(file, worksheet, hdr, schema, tableName, sqlConnStr, loadType);
        }
        else
        {
          // if multiple files, after the first one, the rest need to be appended
          ret = Excel2Table(file, worksheet, hdr, schema, tableName, sqlConnStr, "APPEND");
        }
        
        LogEnd(logDbConnStr, loadDataOutKey, DateTime.Now, (ret.errMsg == SUCCESS ? SUCCESS : FAILURE), ret.rowsLoaded, (ret.errMsg == SUCCESS ? null : ret.errMsg));
        if (ret.retCode != 0) break;
        i++;
      }
      
      return(ret);
    
    } // ExceFilesInDir2Table()
    
    public static void MSSQLGenCreateTableFromQuery(String srcConnStr, String destConnStr, String query, String schemaName, String tableName)
    {
      int count;
      StringBuilder sb = new StringBuilder();
      SqlConnection dstConn = new SqlConnection(destConnStr);
      
      Console.WriteLine("In MSSQLGenCreateTableFromQuery() - check if table needs to be created in destination DB");
      dstConn.Open();
      
      // Does table exist already?
      SqlCommand cmdDst = new SqlCommand();
      cmdDst.CommandTimeout = 0;
      cmdDst.Connection = dstConn;
      cmdDst.CommandText = 
        "select " +
        "  count(*) " +
        "from information_schema.tables " +
        "where table_schema = '" + schemaName + "'" +
        "  and table_name = '" + tableName + "'"
      ;
      count = System.Convert.ToInt32(cmdDst.ExecuteScalar());
            
      if (count != 0) //table exists already?
      {
        dstConn.Close();
        return;
      }
      else
      {
        Console.WriteLine("Table = {0} does not exist, create in destination database.", schemaName + "." + tableName);
      }

      SqlConnection srcConn = new SqlConnection(srcConnStr);
      srcConn.Open();
      SqlCommand cmdSrc = new SqlCommand();
      cmdSrc.CommandTimeout = 0;
      cmdSrc.Connection = srcConn;
      query = query.Replace("select", "select top 1 ");
      cmdSrc.CommandText = query;
      
      try
      {
        SqlDataReader dataReader = cmdSrc.ExecuteReader();
        int fieldCount = dataReader.FieldCount;

        String fieldSep = null;

        Console.WriteLine("Execute SQL = " + query);
        Console.WriteLine("Number of columns in select stmt = " + fieldCount);

        sb.Append("create table " + schemaName + "." + tableName);
        sb.Append("(");

        fieldSep = "  ";
        using (var schemaTable = dataReader.GetSchemaTable())
        {
          foreach (DataRow row in schemaTable.Rows)
          {
            string ColumnName      = row.Field<string>("ColumnName");
            string DataTypeName    = row.Field<string>("DataTypeName");
            short NumericPrecision = row.Field<short>("NumericPrecision");
            short NumericScale     = row.Field<short>("NumericScale");
            int ColumnSize         = row.Field<int>("ColumnSize");
            Console.WriteLine("Column: {0} Type: {1} Precision: {2} Scale: {3} ColumnSize {4}", ColumnName, DataTypeName, NumericPrecision, NumericScale, ColumnSize);
            
            if (DataTypeName.Equals("date") || DataTypeName.Equals("datetime") || DataTypeName.Equals("datetime2") || DataTypeName.Equals("int") || DataTypeName.Equals("bigint") || DataTypeName.Equals("tinyint"))
            {
              sb.Append(fieldSep + ColumnName + " " + DataTypeName);
              fieldSep = ", ";
            }
            else if (DataTypeName.Equals("decimal"))
            {
                sb.Append(fieldSep + ColumnName + " numeric(38, 4)");
                fieldSep = ", ";
            }
            else if (DataTypeName.Equals("char"))
            {
              if (ColumnSize == 1)
              {
                sb.Append(fieldSep + ColumnName + " " + DataTypeName + "(1)");
              }
              else
              {
                sb.Append(fieldSep + ColumnName + " " + DataTypeName + "(" + (ColumnSize > NumericPrecision ? ColumnSize : NumericPrecision) + ")");
              }
              fieldSep = ", ";
            }
            else if (DataTypeName.Equals("varchar") || DataTypeName.Equals("nvarchar"))
            {
               if (ColumnSize > 8000)
               {
                 sb.Append(fieldSep + ColumnName + " " + DataTypeName + "(max)");
               }
               else
               {
                 //sb.Append(fieldSep + ColumnName + " " + DataTypeName + "(8000)");
                 sb.Append(fieldSep + ColumnName + " " + DataTypeName + "(" + ColumnSize + ")");
               }
               fieldSep = ", ";
            }
            else
            {
               Console.WriteLine("Unknown datatype = {0}", DataTypeName);
               sb.Append(fieldSep + ColumnName + " UnknownDataType");
               fieldSep = ", ";
            }
            
          }
        }
        sb.Append(")");
        sb.Append(";");
      }
      catch(Exception e)
      {
        Console.WriteLine(e.ToString());
      }
      finally
      {
        srcConn.Close();
        cmdDst.CommandText = sb.ToString();
        cmdDst.ExecuteNonQuery();
        dstConn.Close();
      }

    } // MSSQLGenCreateTableFromQuery()

    public static void ODBCGenCreateTableFromQuery(String srcConnStr, String destConnStr, String query, String schemaName, String tableName)
    {
      int count;
      StringBuilder sb = new StringBuilder();
      SqlConnection dstConn = new SqlConnection(destConnStr);
      
      Console.WriteLine("In ODBCGenCreateTableFromQuery() - check if table needs to be created in destination DB");
      dstConn.Open();
      
      // Does table exist already?
      SqlCommand cmdDst = new SqlCommand();
      cmdDst.CommandTimeout = 0;
      cmdDst.Connection = dstConn;
      cmdDst.CommandText = 
        "select " +
        "  count(*) " +
        "from information_schema.tables " +
        "where table_schema = '" + schemaName + "'" +
        "  and table_name = '" + tableName + "'"
      ;
      count = System.Convert.ToInt32(cmdDst.ExecuteScalar());
            
      if (count != 0) //table exists already?
      {
        dstConn.Close();
        return;
      }
      else
      {
        Console.WriteLine("Table = {0} does not exist, create in destination database.", schemaName + "." + tableName);
      }

      OdbcConnection srcConn = new OdbcConnection(srcConnStr);
      srcConn.Open();
      OdbcCommand cmdSrc = new OdbcCommand();
      cmdSrc.CommandTimeout = 0;
      cmdSrc.Connection = srcConn;
      query = query.Replace("select", "select top 1 ");
      cmdSrc.CommandText = query;
      
      try
      {
        OdbcDataReader dataReader = cmdSrc.ExecuteReader();
        int fieldCount = dataReader.FieldCount;

        String fieldSep = null;

        Console.WriteLine("Execute SQL = " + query);
        Console.WriteLine("Number of columns in select stmt = " + fieldCount);

        sb.Append("create table " + schemaName + "." + tableName);
        sb.Append("(");

        fieldSep = "  ";
        using (var schemaTable = dataReader.GetSchemaTable())
        {
          foreach (DataRow row in schemaTable.Rows)
          {
            string ColumnName      = row.Field<string>("ColumnName");
            string DataTypeName      = row["DataType"].ToString();
            short NumericPrecision = row.Field<short>("NumericPrecision");
            short NumericScale     = row.Field<short>("NumericScale");
            int ColumnSize         = row.Field<int>("ColumnSize");
           
            Console.WriteLine("Column: {0} Type: {1} Precision: {2} Scale: {3} ColumnSize {4}", ColumnName, DataTypeName, NumericPrecision, NumericScale, ColumnSize);
            
            if (DataTypeName.Equals("System.Date") || DataTypeName.Equals("System.DateTime") || DataTypeName.Equals("System.DateTime2"))
            {
              sb.Append(fieldSep + ColumnName + " " + DataTypeName.Replace("System.", ""));
              fieldSep = ", ";
            }
            else if (DataTypeName.Equals("System.Int16") || DataTypeName.Equals("System.Int32") || DataTypeName.Equals("System.UInt16") || DataTypeName.Equals("System.UInt32") || DataTypeName.Equals("System.SByte"))
            {
              sb.Append(fieldSep + ColumnName + " int");
              fieldSep = ", ";
            }
            
            else if (DataTypeName.Equals("System.Int64") || DataTypeName.Equals("System.UInt64"))
            {
              sb.Append(fieldSep + ColumnName + " bigint");
              fieldSep = ", ";
            }
           
            else if (DataTypeName.Equals("System.Single") || DataTypeName.Equals("System.Decimal") || DataTypeName.Equals("VarNumeric"))
            {
                sb.Append(fieldSep + ColumnName + " numeric(38, 4)");
                fieldSep = ", ";
            }
            else if (DataTypeName.Equals("System.StringFixedLength") || DataTypeName.Equals("AnsiStringFixedLength"))
            {
              if (ColumnSize == 1)
              {
                sb.Append(fieldSep + ColumnName + " char(1)");
              }
              else
              {
                sb.Append(fieldSep + ColumnName + " char(" + (ColumnSize > NumericPrecision ? ColumnSize : NumericPrecision) + ")");
              }
              fieldSep = ", ";
            }
            else if (DataTypeName.Equals("System.String") || DataTypeName.Equals("AnsiString"))
            {
               if (ColumnSize > 8000)
               {
                 sb.Append(fieldSep + ColumnName + " varchar(max)");
               }
               else
               {
                 //sb.Append(fieldSep + ColumnName + " varchar(8000)");
                 sb.Append(fieldSep + ColumnName + " varchar(" + ColumnSize + ")");
               }
               fieldSep = ", ";
            }
            else
            {
               Console.WriteLine("Unknown datatype = {0}", DataTypeName);
               sb.Append(fieldSep + ColumnName + " UnknownDataType");
               fieldSep = ", ";
            }
            
          }
        }
        sb.Append(")");
        sb.Append(";");
      }
      catch(Exception e)
      {
        Console.WriteLine(e.ToString());
      }
      finally
      {
        srcConn.Close();
        cmdDst.CommandText = sb.ToString();
        cmdDst.ExecuteNonQuery();
        dstConn.Close();
      }

    } // ODBCGenCreateTableFromQuery()

    public static void OleDbGenCreateTableFromQuery(String srcConnStr, String destConnStr, String query, String schemaName, String tableName)
    {
      int count;
      StringBuilder sb = new StringBuilder();
      SqlConnection dstConn = new SqlConnection(destConnStr);
      
      Console.WriteLine("In OleDBGenCreateTableFromQuery() - check if table needs to be created in destination DB");
      dstConn.Open();
      
      // Does table exist already?
      SqlCommand cmdDst = new SqlCommand();
      cmdDst.CommandTimeout = 0;
      cmdDst.Connection = dstConn;
      cmdDst.CommandText = 
        "select " +
        "  count(*) " +
        "from information_schema.tables " +
        "where table_schema = '" + schemaName + "'" +
        "  and table_name = '" + tableName + "'"
      ;
      count = System.Convert.ToInt32(cmdDst.ExecuteScalar());
            
      if (count != 0) //table exists already?
      {
        dstConn.Close();
        return;
      }
      else
      {
        Console.WriteLine("Table = {0} does not exist, create in destination database.", schemaName + "." + tableName);
      }

      OleDbConnection srcConn = new OleDbConnection(srcConnStr);
      srcConn.Open();
      OleDbCommand cmdSrc = new OleDbCommand();
      cmdSrc.CommandTimeout = 0;
      cmdSrc.Connection = srcConn;
      query = query.Replace("select", "select top 1 ");
      cmdSrc.CommandText = query;
      
      try
      {
        OleDbDataReader dataReader = cmdSrc.ExecuteReader();
        int fieldCount = dataReader.FieldCount;

        String fieldSep = null;

        Console.WriteLine("Execute SQL = " + query);
        Console.WriteLine("Number of columns in select stmt = " + fieldCount);

        sb.Append("create table " + schemaName + "." + tableName);
        sb.Append("(");

        fieldSep = "  ";
        using (var schemaTable = dataReader.GetSchemaTable())
        {
          foreach (DataRow row in schemaTable.Rows)
          {
            string ColumnName      = row.Field<string>("ColumnName");
            string DataTypeName      = row["DataType"].ToString();
            short NumericPrecision = row.Field<short>("NumericPrecision");
            short NumericScale     = row.Field<short>("NumericScale");
            int ColumnSize         = row.Field<int>("ColumnSize");
           
            Console.WriteLine("Column: {0} Type: {1} Precision: {2} Scale: {3} ColumnSize {4}", ColumnName, DataTypeName, NumericPrecision, NumericScale, ColumnSize);
            
            if (DataTypeName.Equals("System.Date") || DataTypeName.Equals("System.DateTime") || DataTypeName.Equals("System.DateTime2"))
            {
              sb.Append(fieldSep + ColumnName + " " + DataTypeName.Replace("System.", ""));
              fieldSep = ", ";
            }
            else if (DataTypeName.Equals("System.Int16") || DataTypeName.Equals("System.Int32") || DataTypeName.Equals("System.UInt16") || DataTypeName.Equals("System.UInt32") || DataTypeName.Equals("System.SByte"))
            {
              sb.Append(fieldSep + ColumnName + " int");
              fieldSep = ", ";
            }
            
            else if (DataTypeName.Equals("System.Int64") || DataTypeName.Equals("System.UInt64"))
            {
              sb.Append(fieldSep + ColumnName + " bigint");
              fieldSep = ", ";
            }
           
            else if (DataTypeName.Equals("System.Single") || DataTypeName.Equals("System.Decimal") || DataTypeName.Equals("VarNumeric"))
            {
                sb.Append(fieldSep + ColumnName + " numeric(38, 4)");
                fieldSep = ", ";
            }
            else if (DataTypeName.Equals("System.StringFixedLength") || DataTypeName.Equals("AnsiStringFixedLength"))
            {
              if (ColumnSize == 1)
              {
                sb.Append(fieldSep + ColumnName + " char(1)");
              }
              else
              {
                sb.Append(fieldSep + ColumnName + " char(" + (ColumnSize > NumericPrecision ? ColumnSize : NumericPrecision) + ")");
              }
              fieldSep = ", ";
            }
            else if (DataTypeName.Equals("System.String") || DataTypeName.Equals("AnsiString"))
            {
               if (ColumnSize > 8000)
               {
                 sb.Append(fieldSep + ColumnName + " varchar(max)");
               }
               else
               {
                 //sb.Append(fieldSep + ColumnName + " varchar(8000)");
                 sb.Append(fieldSep + ColumnName + " varchar(" + ColumnSize + ")");
               }
               fieldSep = ", ";
            }
            else
            {
               Console.WriteLine("Unknown datatype = {0}", DataTypeName);
               sb.Append(fieldSep + ColumnName + " UnknownDataType");
               fieldSep = ", ";
            }
            
          }
        }
        sb.Append(")");
        sb.Append(";");
      }
      catch(Exception e)
      {
        Console.WriteLine(e.ToString());
      }
      finally
      {
        srcConn.Close();
        cmdDst.CommandText = sb.ToString();
        cmdDst.ExecuteNonQuery();
        dstConn.Close();
      }

    } // OleDBGenCreateTableFromQuery()

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
  
  public static void DropTable(String schema, String tableName, String connectString)
  {
    
    if ( !TableExists(schema, tableName, connectString) ) return;
    
    Console.WriteLine("Dropping destination table {0}.{1}", schema, tableName);

    String sqlStmt = "drop table " + schema + "." + tableName;    
    SqlCommand cmd = new SqlCommand();
    SqlConnection conn = new SqlConnection(connectString);
             
    conn.Open();
    cmd.Connection = conn;
    cmd.CommandText = sqlStmt;
    cmd.ExecuteNonQuery();
    conn.Close();
  
  } // DropTable()
  
  public static void RenameTable(String schema, String tableName, String newTableName, String connectString)
  {
    
    if ( !TableExists(schema, tableName, connectString) ) return;
    
    Console.WriteLine("Rename destination table {0}.{1} to {2}.{3}", schema, tableName, schema, newTableName);

    String sqlStmt = "EXEC sp_rename '" + schema + "." + tableName + "', '" + newTableName + "'";
    SqlCommand cmd = new SqlCommand();
    SqlConnection conn = new SqlConnection(connectString);
             
    conn.Open();
    cmd.Connection = conn;
    cmd.CommandText = sqlStmt;
    cmd.ExecuteNonQuery();
    conn.Close();
  
  } // RenameTable()
    
  public static void CreateTableFromFile(String schema, String tableName, String connectString, string filename, string delim)
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
        //sql.Append("[" + columnName + "]");
        sql.Append(columnName.Replace(" ", "_"));
        sql.Append(" VARCHAR(8000))");
      }
      else
      {
        //sql.Append("[" + columnName + "]");
        sql.Append(columnName.Replace(" ", "_"));
        sql.Append(" VARCHAR(8000),");
      }
    }
    
    // remove last 2 characters
    sql.Length--;
    sql.Length--;
    sql.Append(")"); // put back the closing parenthesis on the data type
    sql.Append(")"); // final closing parenthesis for create table statement
    
    //Console.WriteLine(sql.ToString());
     
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
  
  } // CreateTableFromFile()
  
  public static void CreateTableFromExcel(String schema, String tableName, String connectString, string excelFile, string excelWorksheet)
  {
    String hdr = "YES";
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

    String ddl = "Create table " + schema + "." + tableName + "\n" + "( \n";

    for (int i =0; i < colName.Count; i++)
    {
      if (i == 0)
        ddl = ddl + "  [" + colName[i] + "]";
      else
        ddl = ddl + ", [" + colName[i] + "]";

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
 
  public void RunLoadDefs()
  {
    long loadDataOutKey;
    long numRecs;
    String logDbConnStr = metadataDbConnStr;
    
    int    loadDataDefKey;
    String loadDataName;
    String sourceConnType;
    String sourceDbConnStr;
    String sourceDbQuery;
    String destConnStr;
    String destDbName;
    String destSchemaName;
    String destTableName;
    String loadType;
    int    batchSize;
    String directory;
    String filePattern;
    String worksheet;
    String header;
    String delimiter;
    String activeFlg;
    String procOrSqlBlk;
    String dropTable;
    
    DbReturnStatus ret = null;
    
    Console.WriteLine("Open connect to {0}", metadataDbConnStr);
    OdbcConnection metadataDbConn = new OdbcConnection(metadataDbConnStr);
    metadataDbConn.Open();
    
    String sqlSel =   
    "select " +
    "  e.LoadDataDefKey " +
    ", e.LoadDataName " +
    ", isnull(e.SourceConnType, '') " +
    ", isnull(e.SourceConnStr, '') " +
    ", isnull(e.SourceQuery, '') " +
    ", e.DestConnStr " +
    ", e.DestDbName " +
    ", e.DestSchemaName " +
    ", e.DestTableName " +
    ", e.LoadType " +    
    ", isnull(e.BatchSize, 0) " +
    ", isnull(e.Directory, '') " +
    ", isnull(e.FilePattern, '') " +
    ", isnull(e.Worksheet, '') " +
    ", isnull(e.Header, '') " +
    ", isnull(e.Delimiter, '') " +
    ", e.ActiveFlg " +
    ", isnull(e.GroupName, '') " +
    ", isnull(e.ProcOrSqlBlk, '') " +
    ", isnull(e.DropDestTableIfExists, '') " +
    "from " +
    "  " + metadataDbName + ".dbo.LoadDataDef e " +
    "where 1 = 1 " +
    "  and ActiveFlg = 'Y' " +
    "  and e.loadDataName = case when'" + this.loadDataName + "' = 'ALL' then e.loadDataName else '" + this.loadDataName + "' end " +
    "  and e.groupName = case when'" + groupName + "' = 'ALL' then e.groupName else '" + groupName + "' end " +
    "order by " +
    " e.LoadDataDefKey "
    ;
    
    OdbcCommand cmd = new OdbcCommand(sqlSel, metadataDbConn);
          
    OdbcDataReader reader = cmd.ExecuteReader();
     
    Console.WriteLine();
    Console.WriteLine("Data Load Definitions");
          
    while (reader.Read())
    {
      Console.WriteLine("Start of while loop");
      
      loadDataDefKey  = reader.GetInt32(0);
      loadDataName    = reader.GetString(1);
      sourceConnType  = reader.GetString(2);
      sourceDbConnStr = reader.GetString(3);
      sourceDbQuery   = reader.GetString(4);
      destConnStr     = reader.GetString(5);
      destDbName      = reader.GetString(6);
      destSchemaName  = reader.GetString(7);
      destTableName   = reader.GetString(8);
      loadType        = reader.GetString(9);
      batchSize       = reader.GetInt32(10);
      directory       = reader.GetString(11);
      filePattern     = reader.GetString(12);
      worksheet       = reader.GetString(13);
      header          = reader.GetString(14);
      delimiter       = reader.GetString(15);
      activeFlg       = reader.GetString(16);
      groupName       = reader.GetString(17);
      procOrSqlBlk    = reader.GetString(18);
      dropTable       = reader.GetString(19);

      Console.WriteLine("------------------------------------------------------------------------");
      Console.WriteLine("LoadDataDefKey        = {0}", loadDataDefKey );
      Console.WriteLine("LoadDataName          = {0}", loadDataName );
      Console.WriteLine("SourceConnType        = {0}", sourceConnType );
      Console.WriteLine("SourceDbConnStr       = {0}", sourceDbConnStr ); 
      Console.WriteLine("SourceDbQuery         = {0}", sourceDbQuery ); 
      Console.WriteLine("DestConnStr           = {0}", destConnStr ); 
      Console.WriteLine("DestDbName            = {0}", destDbName ); 
      Console.WriteLine("DestSchemaName        = {0}", destSchemaName ); 
      Console.WriteLine("DestTableName         = {0}", destTableName ); 
      Console.WriteLine("LoadType              = {0}", loadType ); 
      Console.WriteLine("BatchSize             = {0}", batchSize ); 
      Console.WriteLine("Directory             = {0}", directory ); 
      Console.WriteLine("FilePattern           = {0}", filePattern ); 
      Console.WriteLine("Workheet              = {0}", worksheet ); 
      Console.WriteLine("Header                = {0}", header ); 
      Console.WriteLine("Delimiter             = {0}", delimiter ); 
      Console.WriteLine("ActiveFlg             = {0}", activeFlg ); 
      Console.WriteLine("GroupName             = {0}", groupName ); 
      Console.WriteLine("ProcOrSqlBlk          = {0}", procOrSqlBlk ); 
      Console.WriteLine("DropDestTableIfExists = {0}", dropTable ); 
      Console.WriteLine();
          
      numRecs = 0;
      loadDataOutKey = 0;
      
      try
      {
         //ConnectToDb(sourceConnType, sourceDbConnStr);
         
         if (dropTable.Equals("Y"))
         {
           DropTable(destSchemaName, destTableName, destConnStr);
         }

         if (sourceConnType.ToUpper().Equals("MSSQL"))
         {
           loadDataOutKey = LogBeg(logDbConnStr, PROGRAM_NAME, loadDataName, sourceDbQuery, procOrSqlBlk, DateTime.Now, destConnStr, destDbName, destSchemaName, destTableName, loadType, batchSize, directory, filePattern, worksheet, header, delimiter, "", loadDataDefKey);
           MSSQLGenCreateTableFromQuery(sourceDbConnStr, destConnStr, sourceDbQuery, destSchemaName, destTableName);
           Console.WriteLine("Calling BulkCopyTable ...");
           ret = BulkCopyTable(sourceDbConnStr, sourceDbQuery, destConnStr, destSchemaName, destTableName, batchSize, loadType);
           LogEnd(logDbConnStr, loadDataOutKey, DateTime.Now, (ret.errMsg == SUCCESS ? SUCCESS : FAILURE), ret.rowsLoaded, (ret.errMsg == SUCCESS ? null : ret.errMsg));
         }
         else if (sourceConnType.ToUpper().Equals("ODBC"))
         {
           loadDataOutKey = LogBeg(logDbConnStr, PROGRAM_NAME, loadDataName, sourceDbQuery, procOrSqlBlk, DateTime.Now, destConnStr, destDbName, destSchemaName, destTableName, loadType, batchSize, directory, filePattern, worksheet, header, delimiter, "", loadDataDefKey);
           ODBCGenCreateTableFromQuery(sourceDbConnStr, destConnStr, sourceDbQuery, destSchemaName, destTableName);
           Console.WriteLine("Calling OdbcQuery2Table ...");
           ret = OdbcQuery2SqlTable(sourceDbConnStr, sourceDbQuery, destConnStr, destSchemaName, destTableName, loadType);
           LogEnd(logDbConnStr, loadDataOutKey, DateTime.Now, (ret.errMsg == SUCCESS ? SUCCESS : FAILURE), ret.rowsLoaded, (ret.errMsg == SUCCESS ? null : ret.errMsg));
         }
         else if (sourceConnType.ToUpper().Equals("OLEDB"))
         {
           loadDataOutKey = LogBeg(logDbConnStr, PROGRAM_NAME, loadDataName, sourceDbQuery, procOrSqlBlk, DateTime.Now, destConnStr, destDbName, destSchemaName, destTableName, loadType, batchSize, directory, filePattern, worksheet, header, delimiter, "", loadDataDefKey);
           OleDbGenCreateTableFromQuery(sourceDbConnStr, destConnStr, sourceDbQuery, destSchemaName, destTableName);
           Console.WriteLine("Calling OleDbQuery2Table ...");
           ret = OleDbQuery2SqlTable(sourceDbConnStr, sourceDbQuery, destConnStr, destSchemaName, destTableName, loadType);
           LogEnd(logDbConnStr, loadDataOutKey, DateTime.Now, (ret.errMsg == SUCCESS ? SUCCESS : FAILURE), ret.rowsLoaded, (ret.errMsg == SUCCESS ? null : ret.errMsg));
         }
         else if (sourceConnType.ToUpper().Equals("CSV"))
         {
           Console.WriteLine("Calling FilesInDir2Table for CSV file(s) ...");
           ret = FilesInDir2Table(directory, filePattern, header, destSchemaName, destTableName, destConnStr, loadType, ",", sourceConnType,
           logDbConnStr, PROGRAM_NAME, loadDataName, sourceDbQuery, procOrSqlBlk, destDbName, batchSize, worksheet, loadDataDefKey);
         }
         else if (sourceConnType.ToUpper().Equals("FLATFILE"))
         {
           Console.WriteLine("Calling FilesInDir2Table for Flat file(s) ...");
           ret = FilesInDir2Table(directory, filePattern, header, destSchemaName, destTableName, destConnStr, loadType, delimiter, sourceConnType,
           logDbConnStr, PROGRAM_NAME, loadDataName, sourceDbQuery, procOrSqlBlk, destDbName, batchSize, worksheet, loadDataDefKey);
         }
         else if (sourceConnType.ToUpper().Equals("EXCEL"))
         {
           Console.WriteLine("Calling ExcelFilesInDir2Table for Flat file(s) ...");
           ret = ExcelFilesInDir2Table(directory, filePattern, (header.ToUpper() == "Y" ? "YES" : "NO"), destSchemaName, destTableName, destConnStr, loadType, delimiter, sourceConnType,
           logDbConnStr, PROGRAM_NAME, loadDataName, sourceDbQuery, procOrSqlBlk, destDbName, batchSize, worksheet, loadDataDefKey);
         }
         else
         {
           Console.WriteLine("Unknown source system connection type");
           return;
         }
         
         if (procOrSqlBlk != "")
         {
           Console.WriteLine("Execute Procedure or SQL Blk = {0}", procOrSqlBlk);
           cmd = new OdbcCommand(procOrSqlBlk, metadataDbConn);
           cmd.ExecuteNonQuery();
         }

       }
      catch(Exception e)
      {
        String err = e.ToString();
        Console.WriteLine("ERROR! in RunLoadDefs()");
        Console.WriteLine(err);
        LogEnd(logDbConnStr, loadDataOutKey, DateTime.Now, "FAILURE", numRecs, err);
      }
      finally
      {
        //DisconnectFromDb(sourceConnType);
        Console.WriteLine("End of running extracts");
      }
       
    }
    
    reader.Close();
    metadataDbConn.Close();
    metadataDbConn.Dispose();

  } // RunLoadDefs()
  
  public static long LogBeg(String connStr, String programName, String loadDataName, String sqlStatement, String procOrSqlBlk, DateTime begDateTime, String destConnStr, String destDbName, String destSchemaName, String destTableName, String loadType, int batchSize, String directory, String filePattern, String worksheet, String header, String delimiter, String filename, long loadDataDefKey)
  {
    String dbName = metadataDbName;
    
    OdbcConnection odbcConn = new OdbcConnection(connStr);
    odbcConn.Open();
    
    OdbcCommand cmd = new OdbcCommand();
    cmd.Connection = odbcConn;
    cmd.CommandText = "{call " + dbName + ".dbo.LogBegLoadData(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)}";
    cmd.CommandType = CommandType.StoredProcedure;
    cmd.Parameters.AddWithValue("@ProgramName", programName);
    cmd.Parameters.AddWithValue("@LoadDataName", loadDataName);
    cmd.Parameters.AddWithValue("@SqlStatement", sqlStatement);
    cmd.Parameters.AddWithValue("@ProcOrSqlBlk", procOrSqlBlk);
    //cmd.Parameters.AddWithValue("@BegDateTime", begDateTime);
    cmd.Parameters.AddWithValue("@BegDateTime", begDateTime.ToString("yyyy-MM-dd HH:mm:ss")); // work around for odbc fractional seconds exception thrown
    cmd.Parameters.AddWithValue("@DestConnStr", destConnStr);
    cmd.Parameters.AddWithValue("@DestDbName", destDbName);
    cmd.Parameters.AddWithValue("@DestSchemaName", destSchemaName);
    cmd.Parameters.AddWithValue("@DestTableName", destTableName);
    cmd.Parameters.AddWithValue("@LoadType", loadType);
    cmd.Parameters.AddWithValue("@BatchSize", batchSize);
    cmd.Parameters.AddWithValue("@Directory", directory);
    cmd.Parameters.AddWithValue("@FilePattern", filePattern);
    cmd.Parameters.AddWithValue("@Worksheet", worksheet);
    cmd.Parameters.AddWithValue("@Header", header);
    cmd.Parameters.AddWithValue("@Delimiter", delimiter);
    cmd.Parameters.AddWithValue("@Filename", filename);
    cmd.Parameters.AddWithValue("@LoadDataDefKey", loadDataDefKey);
    
    long dataLoadOutKey = Convert.ToInt64(cmd.ExecuteScalar());
    odbcConn.Close();
    
    return dataLoadOutKey;
        
  } // LogBeg()
  
  public static void LogEnd(String connStr, long loadDataOutKey, DateTime endDateTime, String status, long numRecs, String errorMsg = null)
  {
    String dbName = metadataDbName;
    
    OdbcConnection odbcConn = new OdbcConnection(connStr);
    odbcConn.Open();
  
    OdbcCommand cmd = new OdbcCommand();
    cmd.Connection = odbcConn;
    
    if (errorMsg == null)
    {
      cmd.CommandText = "{call " + dbName + ".dbo.LogEndLoadData(?, ?, ?, ?)}";
    }
    else
    {
      cmd.CommandText = "{call " + dbName + ".dbo.LogEndLoadData(?, ?, ?, ?, ?)}";
    }
    cmd.CommandType = CommandType.StoredProcedure;
    
    cmd.Parameters.AddWithValue("@LoadDataKeyOutKey", loadDataOutKey);
    cmd.Parameters.AddWithValue("@EndDateTime", endDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
    cmd.Parameters.AddWithValue("@status", status);
    cmd.Parameters.AddWithValue("@NumRecs", numRecs);
        
    if (errorMsg != null)
    {
      cmd.Parameters.AddWithValue("@ErrorMsg", errorMsg);
    }
    cmd.ExecuteNonQuery();
    odbcConn.Close();
    
  } // LogEnd()
  
  public static void Usage()
  {
    Console.WriteLine();
    Console.WriteLine("--------------------------------------------------------------------------------------------------");
    Console.WriteLine("Usage: {0} arg1 arg3 arg3 arg4", PROGRAM_NAME);
    Console.WriteLine();
    Console.WriteLine("Examples:");
    Console.WriteLine();
    Console.WriteLine(">{0} metadataODBConnStr dbName loadDataName groupName", PROGRAM_NAME);
    Console.WriteLine();
    Console.WriteLine();
    Console.WriteLine("  metadataODBConnStr - Metadata database connection string where table of load definitions exists");    
    Console.WriteLine("  dbName             - Database name where table of load definitions exists");
    Console.WriteLine("  loadDataName       - Data Load Name or specify ALL to run everyone of them");
    Console.WriteLine("  groupName          - Group Name or specify ALL to run every group");
    Console.WriteLine();
    Console.WriteLine("     Connection examples:");
    Console.WriteLine();
    Console.WriteLine("     --> \"Driver={SQL Server};Server=theServer;Database=theDatabase;Trusted_Connection=yes;\"");
    Console.WriteLine("     --> \"Driver={SQL Server};Server=theServer;Database=theDatabase;uid=username;pwd=password;\"");
    Console.WriteLine();
    Console.WriteLine("--------------------------------------------------------------------------------------------------");
    Console.WriteLine();    
  }

  /*
   * Main() - Entry point of program.
   */
  public static void Main (String [] args)
  {
    String action;
    
    String metadataDbConnStr = null;
    String metadataDbName    = null;
    String loadDataName     = null;
    String groupName         = null;
   
    LoadData ld = null;
   
    Console.WriteLine("Starting Program{0}, number of command line arguments = {1}", PROGRAM_NAME, args.Length);
    
    if (args.Length != 4)
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
    
    metadataDbConnStr  = args[0];
    metadataDbName     = args[1];
    loadDataName       = args[2];
    groupName          = args[3];
          
    ld = new LoadData(metadataDbConnStr, metadataDbName, loadDataName, groupName);
      
    try
    {
       ld.RunLoadDefs();
    }
    catch(Exception e)
    {
      Console.WriteLine("ERROR!");
      Console.WriteLine(e.ToString());
    }
    finally
    {
       Console.WriteLine("Done!");
    }
   
  } // Main()

} // LoadData class

} // Data namespace
