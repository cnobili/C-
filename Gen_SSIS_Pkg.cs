/*
 * Gen_SSIS_Pkg.cs
 *
 * Generates a SSIS Package to load a flat file into SQL database table.
 *
 * Craig Nobili
 *
 * Need to link in the following DLLs when compiling:
 *
 *   Microsoft.SqlServer.DTSPipelineWrap.dll
 *   Microsoft.SqlServer.DTSRuntimeWrap.dll
 *   Microsoft.SqlServer.ManagedDTS.dll
 *   Microsoft.SqlServer.SQLTaskConnectionsWrap.dll
 *
 */
using System;
using System.IO;
using System.Text;
using System.Data.SqlClient;
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;
using Wrapper = Microsoft.SqlServer.Dts.Runtime.Wrapper;
using Microsoft.SqlServer.Dts.Tasks.ExecuteSQLTask;

namespace FileLoader
{

class GenSSISpkg
{

  static void Main(string[] args)
  {
    Console.Title = "File Loader";
    Console.ForegroundColor = ConsoleColor.Yellow;

    // Read input parameters
    if (args.Length != 4)
    {
      Console.WriteLine("\nUsage: Gen_SSIS_Pkg flatfile delimiter server database");
      Console.WriteLine();
      Console.WriteLine("  flatfile  - Full Path of flat file (should have a header record in it)");
      Console.WriteLine("  delimiter - Field Delimiter");
      Console.WriteLine("  server    - Server");
      Console.WriteLine("  database  - Database");
      Console.WriteLine();
      Console.WriteLine("Example> Gen_SSIS_Pkg c:\\mydir\\myfile.dat , TheServer TheDatabase");
      Console.WriteLine();
      Console.WriteLine("The table will be created with the same name as the flatfile, without");
      Console.WriteLine("the extension. If a table by this name exists in the database already,");
      Console.WriteLine("you must drop it first or just change the flatfile name.");
      Console.WriteLine();
      Console.WriteLine("The generated package can also be used to load the file.");
      Console.WriteLine("You may want to truncate the table first.");
      Console.WriteLine();
      Console.WriteLine("Example> dtexec /f Load_filename.dtsx");
      return;
    }
    var file = args[0];
    var delim = args[1];
    var server = args[2];
    var database = args[3];

    Console.WriteLine("Building Package...");

    // Create a new SSIS Package
    var package = new Package();
    
    // Add a Connection Manager to the Package, of type, FLATFILE
    var connMgrFlatFile = package.Connections.Add("FLATFILE");
    Console.WriteLine("Added FLATFILE Connection Manager");

    connMgrFlatFile.ConnectionString = file;
    connMgrFlatFile.Name = "My Import File Connection";
    connMgrFlatFile.Description = "Flat File Connection";
    
    // Get the Column names to be used in configuring the Flat File Connection 
    // by reading the first line of the Import File which contains the Field names
    string[] columns = null;

    using (var stream = new StreamReader(file))
    {
      var fieldNames = stream.ReadLine();
      //if (fieldNames != null) columns = fieldNames.Split("\t".ToCharArray());
      if (fieldNames != null) columns = fieldNames.Split(delim.ToCharArray());
    }

    // Configure Columns and their Properties for the Flat File Connection Manager
    var connMgrFlatFileInnerObj = (Wrapper.IDTSConnectionManagerFlatFile100)connMgrFlatFile.InnerObject;

    connMgrFlatFileInnerObj.RowDelimiter = "\r\n";
    connMgrFlatFileInnerObj.ColumnNamesInFirstDataRow = true;

    if (columns == null)
    {
      Console.WriteLine("The flat file {0} must contain a header record with the columns", file);
      return;
    }
    
    foreach (var column in columns)
    {
      // Add a new Column to the Flat File Connection Manager
      var flatFileColumn = connMgrFlatFileInnerObj.Columns.Add();

      flatFileColumn.DataType = Wrapper.DataType.DT_WSTR;
      flatFileColumn.ColumnWidth = 255;

      //flatFileColumn.ColumnDelimiter = columns.GetUpperBound(0) == Array.IndexOf(columns, column) ? "\r\n" : "\t";
      flatFileColumn.ColumnDelimiter = columns.GetUpperBound(0) == Array.IndexOf(columns, column) ? "\r\n" : delim;
      flatFileColumn.ColumnType = "Delimited";

      // Use the Import File Field name to name the Column
      var columnName = flatFileColumn as Wrapper.IDTSName100;
      if (columnName != null) columnName.Name = column;
    }

    // Add a Connection Manager to the Package, of type, OLEDB 
    var connMgrOleDb = package.Connections.Add("OLEDB");
    Console.WriteLine("Adding OLEDB Destination Connection Manager");

    var connectionString = new StringBuilder();

    connectionString.Append("Provider=SQLOLEDB.1;");
    connectionString.Append("Integrated Security=SSPI;Initial Catalog=");
    connectionString.Append(database);
    connectionString.Append(";Data Source=");
    connectionString.Append(server);
    connectionString.Append(";");

    connMgrOleDb.ConnectionString = connectionString.ToString();
    connMgrOleDb.Name = "My OLE DB Connection";
    connMgrOleDb.Description = "OLE DB connection";

    // Add a Data Flow Task to the Package
    Console.WriteLine("Adding Data Flow Task to Package");
    var e = package.Executables.Add("STOCK:PipelineTask");
    var mainPipe = e as TaskHost;

    if (mainPipe == null)
    {
      Console.WriteLine("Error creating Data Flow Task");
      package.Dispose();
      return;
    }
    
    mainPipe.Name = "MyDataFlowTask";
    var dataFlowTask = mainPipe.InnerObject as MainPipe;

    var app = new Application();

    if (dataFlowTask == null)
    {
      Console.WriteLine("Error creating Data Flow Task");
      package.Dispose();
      return;
    }
    
    // Add a Flat File Source Component to the Data Flow Task
    Console.WriteLine("Adding flat file source component to data flow task");
    var flatFileSourceComponent = dataFlowTask.ComponentMetaDataCollection.New();
    flatFileSourceComponent.Name = "My Flat File Source";
    flatFileSourceComponent.ComponentClassID = app.PipelineComponentInfos["Flat File Source"].CreationName;

    // Get the design time instance of the Flat File Source Component
    var flatFileSourceInstance = flatFileSourceComponent.Instantiate();
    flatFileSourceInstance.ProvideComponentProperties();

    flatFileSourceComponent.RuntimeConnectionCollection[0].ConnectionManager =
    DtsConvert.GetExtendedInterface(connMgrFlatFile);

    flatFileSourceComponent.RuntimeConnectionCollection[0].ConnectionManagerID = connMgrFlatFile.ID;

    // Reinitialize the metadata.
    flatFileSourceInstance.AcquireConnections(null);
    flatFileSourceInstance.ReinitializeMetaData();
    flatFileSourceInstance.ReleaseConnections();
    
    // Add an OLE DB Destination Component to the Data Flow
    Console.WriteLine("Adding flat OLE DB Destination component to data flow task");
    var oleDbDestinationComponent = dataFlowTask.ComponentMetaDataCollection.New();
    oleDbDestinationComponent.Name = "MyOLEDBDestination";
    oleDbDestinationComponent.ComponentClassID = app.PipelineComponentInfos["OLE DB Destination"].CreationName;

    // Get the design time instance of the Ole Db Destination component
    var oleDbDestinationInstance = oleDbDestinationComponent.Instantiate();
    oleDbDestinationInstance.ProvideComponentProperties();

    // Set Ole Db Destination Connection
    oleDbDestinationComponent.RuntimeConnectionCollection[0].ConnectionManagerID = connMgrOleDb.ID;
    oleDbDestinationComponent.RuntimeConnectionCollection[0].ConnectionManager =
    DtsConvert.GetExtendedInterface(connMgrOleDb);

    // Set destination load type
    Console.WriteLine("Setting the destination load type");
    oleDbDestinationInstance.SetComponentProperty("AccessMode", 3);

    // Create table in destination sql database to hold file data
    var sql = new StringBuilder();

    sql.Append("CREATE TABLE ");
    sql.Append(Path.GetFileNameWithoutExtension(file));
    sql.Append(" ( ");

    foreach (var columnName in columns)
    {
      if (columns.GetUpperBound(0) == Array.IndexOf(columns, columnName))
      {
        sql.Append(columnName);
        sql.Append(" NVARCHAR(4000))");
      }
      else
      {
        sql.Append(columnName);
        sql.Append(" NVARCHAR(4000),");
      }
    }
   
    var connection = new SqlConnection(
    string.Format("Data Source={0};Initial Catalog={1};Integrated Security=TRUE;", server, database));
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
      
    // Now set Ole Db Destination Table name
    oleDbDestinationInstance.SetComponentProperty("OpenRowset", Path.GetFileNameWithoutExtension(file));

    // Create a Precedence Constraint between Flat File Source and OLEDB Destination Components
    var path = dataFlowTask.PathCollection.New();
    path.AttachPathAndPropagateNotifications(flatFileSourceComponent.OutputCollection[0],
    oleDbDestinationComponent.InputCollection[0]);
    Console.WriteLine("Created precedence constraint between flat file source and ole db destination");

    // Get the list of available columns
    var oleDbDestinationInput = oleDbDestinationComponent.InputCollection[0];
    var oleDbDestinationvInput = oleDbDestinationInput.GetVirtualInput();

    var oleDbDestinationVirtualInputColumns = oleDbDestinationvInput.VirtualInputColumnCollection;

    // Reinitialize the metadata
    Console.WriteLine("Reinitialize the metadata");
    oleDbDestinationInstance.AcquireConnections(null);
    oleDbDestinationInstance.ReinitializeMetaData();
    oleDbDestinationInstance.ReleaseConnections();

    // Map Flat File Source Component Output Columns to Ole Db Destination Input Columns
    Console.WriteLine("Mapping flat file source output columns to ole db destination input columns");
    foreach (IDTSVirtualInputColumn100 vColumn in oleDbDestinationVirtualInputColumns)
    {
      var inputColumn = oleDbDestinationInstance.SetUsageType(oleDbDestinationInput.ID,
      oleDbDestinationvInput, vColumn.LineageID, DTSUsageType.UT_READONLY);

      var externalColumn = oleDbDestinationInput.ExternalMetadataColumnCollection[inputColumn.Name];

      oleDbDestinationInstance.MapInputColumn(oleDbDestinationInput.ID, inputColumn.ID, externalColumn.ID);
    }
    
    Console.WriteLine("Executing Package...");
    package.Execute();

    var dtsx = new StringBuilder();
    //dtsx.Append(Path.GetDirectoryName(file)).Append("\\").Append("Load_").Append(Path.GetFileNameWithoutExtension(file)).Append(".dtsx");
    // Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); Path.GetDirectoryName(Application.ExecutablePath);
    dtsx.Append(Directory.GetCurrentDirectory()).Append("\\").Append("Load_").Append(Path.GetFileNameWithoutExtension(file)).Append(".dtsx");

    Console.WriteLine("Saving Package {0}", dtsx.ToString());
    app.SaveToXml(dtsx.ToString(), package, null);
       
    package.Dispose();
    Console.WriteLine("Done");
        
  } // Main()
        
  } // class GenSSISpkg
    
} // namespace FileLoader
