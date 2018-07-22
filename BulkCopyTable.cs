/*
 * BulkCopyTable.cs
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

class BulkCopyTable
{
  private static String srcConnStr;
  private static String dstConnStr;
  private static String srcTable;
  private static String dstTable;
  
  /*
   * CopyTable()
   */
  public static void CopyTable()
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
  
  } // CopyTable()
  
  public static void Main(String[] args)
  {
    
    if (args.Length != 4)
    {
      Console.WriteLine("Usage:BulkCopyTable  <srcConnStr> <srcTable> <dstConnStr> <dstTable>");
      Console.WriteLine();
      Console.WriteLine("Example:\n BulkCopyTable \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" sourceTable \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" destinationTable");
      return;
    }
    srcConnStr = args[0];
    srcTable   = args[1];
    dstConnStr = args[2];
    dstTable   = args[3];
    
    CopyTable();
    
  } // Main()
  
} // BulkCopyTable class

