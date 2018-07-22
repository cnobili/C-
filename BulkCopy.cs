/*
 * BulkCopy.cs
 *
 * Bulk copies data from an external table (any database) to a 
 * corresponding table in SQL Server (has to be SQL for Bulk Operation).
 *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SqlClient;

class BulkCopy
{
  private static String dstServer;
  private static String dstDB;
  private static String dstUser;
  private static String dstPass;
    
  private static String clarityServer;
  private static String clarityDB;
  private static String clarityUser;
  private static String clarityPass;
  
  /*
   * GetConnectionData
   *
   * Gets database connection information from config file.
   */
  public static void GetConnectionData(String configFile)
  {
    dstServer   = ConfigurationManager.AppSettings["dstServer"];
    dstDB       = ConfigurationManager.AppSettings["dstDB"];
    dstUser     = ConfigurationManager.AppSettings["dstUser"];
    dstPass     = ConfigurationManager.AppSettings["dstPass"];
    
    clarityServer   = ConfigurationManager.AppSettings["clarityServer"];
    clarityDB = ConfigurationManager.AppSettings["clarityDB"];
    clarityUser     = ConfigurationManager.AppSettings["clarityUser"];
    clarityPass     = ConfigurationManager.AppSettings["clarityPass"];

  } // GetConnectionData()
  
  /*
   * GetSrcConnStr()
   *
   * Returns Source System Connection String.
   */
  public static String GetSrcConnStr() 
  {
       
    String sqlConn = "user id=" + clarityUser + "; password=" + clarityPass 
      + "; server=" + clarityServer + "; Trusted_Connection=false; database=" + clarityDB;
        
    return(sqlConn);
  
  } // GetSrcConnStr()
  
  /*
   * GetDstConnStr()
   *
   * Returns Destination System Connection String.
   */
  public static String GetDstConnStr() 
  {
    String sqlConn = "user id=" + dstUser + "; password=" + dstPass 
      + "; server=" + dstServer + "; Trusted_Connection=false; database=" + dstDB;
      
   return(sqlConn);
  
  } // GetDstConnDst()
  
  /*
   * CopyTable()
   */
  public static void CopyTable()
  {
    // Open a sourceConnection
    using (SqlConnection sourceConnection = new SqlConnection(GetSrcConnStr()))
    {
      sourceConnection.Open();

      // Perform an initial count
      SqlCommand commandRowCount = new SqlCommand
      (
        "select " +
        "  count(*) " +
        "from " +
        "  cr_stat_extract e with (nolock) " 
      ,  sourceConnection
      );
      
      long countStart = System.Convert.ToInt32(
      commandRowCount.ExecuteScalar());
      Console.WriteLine("Starting row count = {0}", countStart);

      // Get data from the source table as a SqlDataReader.
      SqlCommand commandSourceData = new SqlCommand
      (
        "select " +
        "  e.exec_descriptor " +
        ", e.end_time " +
        ", e.error_code " +
        "from " +
        "  cr_stat_extract e with (nolock) " 
      , sourceConnection
      );
      
      SqlDataReader reader = commandSourceData.ExecuteReader();

      // Open the destination connection
      using (SqlConnection destinationConnection = new SqlConnection(GetDstConnStr()))
      {
        destinationConnection.Open();
        
        // Set up the bulk copy object.  
        // Note that the column positions in the source 
        // data reader match the column positions in  
        // the destination table so there is no need to 
        // map columns. 
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
        {
          bulkCopy.DestinationTableName = "dbo.test_cr_stat_extract";

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
        
        commandRowCount.Connection = destinationConnection;
        commandRowCount.CommandText = 
        "select " +
        "  count(*) " +
        "from " +
        "  test_cr_stat_extract e with (nolock) "
        ;
        long countEnd = System.Convert.ToInt32(commandRowCount.ExecuteScalar());
        Console.WriteLine("Ending row count = {0}", countEnd);
        Console.WriteLine("{0} rows were added.", countEnd - countStart);
        Console.WriteLine("Press Enter to finish.");
        Console.ReadLine();
      }
    }  
  
  } // CopyTable()
  
  public static void Main(String[] args)
  {
    // Get filename on the command line
    if (args.Length != 1)
    {
      Console.WriteLine("Usage:BulkCopy <configFile>");
      return;
    }
    String configFile = args[0];
    
    if (!File.Exists(configFile))
    {
      Console.WriteLine("filename = <{0}> not found", configFile);
      return;
    }
  
    AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", configFile);
    
    GetConnectionData(configFile);
    CopyTable();
    
  } // Main()
  
} // BulkCopy class
