/*
 * Make_SSIS_Pkg_File2Table.cs
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

namespace JMH_Make_SSIS_Gen_File2Table
{

class FileLoader
{

  // Constants
  private const String ACTION_LOAD_FILE           = "LoadFile";
  private const String ACTION_LOAD_DIR            = "LoadDir";
  private const String ACTION_LOAD_FROM_FILE_LIST = "LoadFromFileList";
  
  // Class Variables
  
  public static void Usage()
  {
    Console.WriteLine();
    Console.WriteLine("------------------------------------------------------------------------------------------------");
    Console.WriteLine("Usage: Make_SSIS_Pkg_File2Table arg1 arg3 arg3 ... argN");
    Console.WriteLine();
    Console.WriteLine("  where argument = action");
    Console.WriteLine("  followed by one or more other arguments depending on the action");
    Console.WriteLine();
    Console.WriteLine("action = LoadFile | LoadDir | LoadFromFileList");
    Console.WriteLine();
    Console.WriteLine("Examples:");
    Console.WriteLine();
    Console.WriteLine(">SSIS_Pkg_File2Table LoadFile flatFile delimiter schema tableName truncate|append connectString");
    Console.WriteLine(">SSIS_Pkg_File2Table LoadDir  directory filePattern delimiter schema truncate|append connectString");
    Console.WriteLine(">SSIS_Pkg_File2Table LoadFromFileList listOfFiles");
    Console.WriteLine("------------------------------------------------------------------------------------------------");
    Console.WriteLine();
    Console.WriteLine("Example> SSIS_Pkg_File2Table LoadFile c:\\mydir\\myfile.dat , dbo tableName Truncate \"Integrated Security=SSPI;Initial Catalog=theDatabase;Data Source=theServer;\" ");
    Console.WriteLine();
    Console.WriteLine("The table will be created if it does not exist already.");
    Console.WriteLine();
    Console.WriteLine("The generated package can also be used to load the file (appends).");
    Console.WriteLine();
    Console.WriteLine("Example> dtexec /f Load_filename.dtsx");
  
  } // Usage()
  
  public static long GetFileRecCount(String filename)
  {
    String[] lines = File.ReadAllLines(filename);
    return(lines.Length);
      
  } // GetFileRecCount()
  
  public static String UpperFirstChar(String s)
  {
  
     // Check for empty string.
     if (string.IsNullOrEmpty(s))
       return string.Empty;
          
     return char.ToUpper(s[0]) + s.Substring(1);
  
  } // UpperFirstChar()
  
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
  
  public static void CreateTable(String schema, String tableName, String connectString, string[] columns)
  {
  
    // Create table in destination sql database to hold file data
    var sql = new StringBuilder();
  
    sql.Append("CREATE TABLE ");
    sql.Append(schema + "." + tableName);
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

  public static void TruncateTable(String schema, String tableName, String connectString)
  {
  
    // Truncate destination table
    var sql = new StringBuilder();
  
    sql.Append("truncate table ");
    sql.Append(schema + "." + tableName);
        
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
  
  } // TruncateTable()
 
  public static void LoadFromFileList(String fileList)
  {
    String filename;
    String delim;
    String schema;
    String tableName;
    String loadType;
    String connectString;
    
    StreamReader file;
    String line;
    String[] tok;
    
    if (!File.Exists(fileList))
    {
      Console.WriteLine("File = {0} does not exist", fileList);
      return;
    }
    
    file = new System.IO.StreamReader(fileList);
    
    while((line = file.ReadLine()) != null)
    {
      // Skip blank lines and comments
      if ( line.StartsWith("REM") || line == "") continue;
      
      Console.WriteLine("line = {0}", line);
      
      tok = line.Split('*');
      
      filename      = tok[0];
      delim         = tok[1];
      schema        = tok[2];
      tableName     = tok[3];
      loadType      = tok[4];
      connectString = tok[5];
      
      LoadFile(filename, delim, schema, tableName, loadType, connectString);
      
    }
    
  } // LoadFromConfigFile()
    
  public static void LoadFilesInDir(String dir, String filePattern, String delim, String schema, String loadType, String connectString)
  {
    String tableName;
    
    if ( !Directory.Exists(dir) )
    {
      Console.WriteLine("Error: input directory = {0} does not exist", dir);
      return;
    }
  
    String[] filePaths = Directory.GetFiles(dir, filePattern); 
    
    if (filePaths.Length == 0)
    {
      Console.WriteLine("No files in '{0}' matching filePattern = {1}", dir, filePattern);
      return;
    }
  
    foreach (String fp in filePaths)
    {
      tableName = Path.GetFileNameWithoutExtension(fp);
      Console.WriteLine("Loading file = {0} into table = {1}", fp, tableName);
      LoadFile(fp, delim, schema, tableName, loadType, connectString);
    }
  
  } // LoadFilesInDir()
  
  public static void LoadFile(String filename, String delim, String schema, String tableName, String loadType, String connectString)
  {
  
    if (!File.Exists(filename))
    {
      Console.WriteLine("File = {0} does not exist", filename);
      return;
    }
    
    String identityColumn = UpperFirstChar(tableName) + "ID";
    String filenameWithoutPath = Path.GetFileName(filename);
  
    Console.WriteLine("Building Package...");

    // Create a new SSIS Package
    var package = new Package();
    
    // Add a Connection Manager to the Package, of type, FLATFILE
    var connMgrFlatFile = package.Connections.Add("FLATFILE");
    Console.WriteLine("Added FLATFILE Connection Manager");

    connMgrFlatFile.ConnectionString = filename;
    connMgrFlatFile.Name = "My Import File Connection";
    connMgrFlatFile.Description = "Flat File Connection";
    
    // Get the Column names to be used in configuring the Flat File Connection 
    // by reading the first line of the Import File which contains the Field names
    string[] columns = null;

    using (var stream = new StreamReader(filename))
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
      Console.WriteLine("The flat file {0} must contain a header record with the columns", filename);
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
      //if (columnName != null) columnName.Name = column;
      if (columnName != null) columnName.Name = column.Trim();
    }

    // Add a Connection Manager to the Package, of type, OLEDB 
    var connMgrOleDb = package.Connections.Add("OLEDB");
    Console.WriteLine("Adding OLEDB Destination Connection Manager");
    
    connMgrOleDb.ConnectionString = "Provider=SQLOLEDB.1;" + connectString;
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
      
    // Now set Ole Db Destination Table name
    oleDbDestinationInstance.SetComponentProperty("OpenRowset", schema + "." + tableName);

    // Create a Precedence Constraint between Flat File Source and OLEDB Destination Components
    var path = dataFlowTask.PathCollection.New();
    path.AttachPathAndPropagateNotifications(flatFileSourceComponent.OutputCollection[0],
    oleDbDestinationComponent.InputCollection[0]);
    Console.WriteLine("Created precedence constraint between flat file source and ole db destination");

    // Get the list of available columns
    var oleDbDestinationInput = oleDbDestinationComponent.InputCollection[0];
    var oleDbDestinationvInput = oleDbDestinationInput.GetVirtualInput();

    var oleDbDestinationVirtualInputColumns = oleDbDestinationvInput.VirtualInputColumnCollection;

    // Create Table if does not exist
    if (!TableExists(schema, tableName, connectString))
    {
      Console.WriteLine("Creating table " + tableName);
      CreateTable(schema, tableName, connectString, columns);
    }
    
    if (loadType.ToUpper() == "TRUNCATE")
    { 
      Console.WriteLine("Truncating table = " + tableName);
      TruncateTable(schema, tableName, connectString);
    }
     
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
    //dtsx.Append(Directory.GetCurrentDirectory()).Append("\\").Append("Load_").Append(Path.GetFileNameWithoutExtension(filename)).Append(".dtsx");
    dtsx.Append(Directory.GetCurrentDirectory()).Append("\\").Append("Load_").Append(tableName).Append(".dtsx");

    Console.WriteLine("Saving Package {0}", dtsx.ToString());
    app.SaveToXml(dtsx.ToString(), package, null);
       
    package.Dispose();
    
    Console.WriteLine("Records in file = {0}, header record excluded", GetFileRecCount(filename) - 1);
    Console.WriteLine("Done");
  
  } // LoadFile()
    
  public static void Main(string[] args)
  {
  
    String filename;
    String dir;
    String filePattern;
    String delim;
    String schema;
    String tableName;
    String loadType;
    String connectString;
    String fileList;
    
    Console.Title = "File Loader";
    Console.ForegroundColor = ConsoleColor.Yellow;
   
    if (args.Length == 0)
    {
      Usage();
      return;
    }
    
    Console.WriteLine("\nSSIS_Pkg_File2Table invoked with the following command line arguments:");
    String action = args[0];
    for (int i = 0; i < args.Length; i++)
    {
      Console.WriteLine("  arg {0} = {1}", i + 1, args[i]);
    }
    
    //
    // Call handler routine based on action passed in
    //
    if ( action.ToUpper().Equals(ACTION_LOAD_FILE.ToUpper()) )
    {
      if (args.Length != 7)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      filename        = args[1];
      delim           = args[2];
      schema          = args[3];
      tableName       = args[4];
      loadType        = args[5];
      connectString   = args[6];
      
      LoadFile(filename, delim, schema, tableName, loadType, connectString);

    }
    else if ( action.ToUpper().Equals(ACTION_LOAD_DIR.ToUpper()) )
    {
      if (args.Length != 7)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      dir             = args[1];
      filePattern     = args[2];
      delim           = args[3];
      schema          = args[4];
      loadType        = args[5];
      connectString   = args[6];
            
      LoadFilesInDir(dir, filePattern, delim, schema, loadType, connectString);
      
    }
    else if ( action.ToUpper().Equals(ACTION_LOAD_FROM_FILE_LIST.ToUpper()) )
    {
      if (args.Length != 2)
      {
        Console.WriteLine("\nWrong number of arguments for action = {0}", action);
        Usage();
        return;
      }
      
      fileList        = args[1];
            
      LoadFromFileList(fileList);
      
    }
    else
    {
      Console.WriteLine("\nUnknown action = {0}", action);
      Usage();
      return;
    }
            
  } // Main()
        
  } // class FileLoader
    
} // namespace JMH_Make_SSIS_Gen_File2Table
