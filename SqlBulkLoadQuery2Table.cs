/*
 * SqlBulkLoadQuery2Table.cs
 *
 * Bulk copies table data from SQL Server Source DB to SQL Server Destination DB.
  *
 * Craig Nobili
 */

using System;
using System.Configuration;
using System.IO;
using System.Text;
using System.Data;
using System.Data.SqlClient;

class SqlBulkLoadQuery2Table
{
  private static String srcConnStr;
  private static String dstConnStr;
  private static String srcQueryFile;
  private static String dstTable;
  
   private static String srcQuery;
   
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
   * CopyData()
   */
  public static void CopyData()
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
  
  } // CopyData()
  
  public static void Main(String[] args)
  {
    
    if (args.Length != 4)
    {
      Console.WriteLine("Usage:SqlBulkLoadQuery2Table  <srcConnStr> <srcQueryFile> <dstConnStr> <dstTable>");
      Console.WriteLine();
      Console.WriteLine("Example:\n SqlBulkLoadQuery2Table \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" srcQueryFile.sql \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" destinationTable");
      return;
    }
    srcConnStr   = args[0];
    srcQueryFile = args[1];
    dstConnStr   = args[2];
    dstTable     = args[3];
    
    srcQuery = FileToString(srcQueryFile);
    
    if (srcQuery != null)
    {
      CopyData();
    }
    
  } // Main()
  
} // SqlBulkQuery2Table class

