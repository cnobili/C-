/*
 * Gen_SSIS_Pkg_Query2Table.cs
 *
 * Generates a SSIS Package to load a SQL database destination table based on
 * a source system query.
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

namespace JMH_SSIS_Gen_Query2Table
{

class LoadTableFromQuery
{

  static void Main(string[] args)
  {
    Console.Title = "LoadTableFromQuery";
    Console.ForegroundColor = ConsoleColor.Yellow;

    // Read input parameters
    if ( !(args.Length == 4 || args.Length == 5) )
    {
      Console.WriteLine("\nUsage: Gen_SSIS_Pkg_Query2Table srcConnectStr srcQuery dstConnectStr dstTable [truncate]");
      Console.WriteLine();
      Console.WriteLine("  srcConnectStr - Source database connection string");
      Console.WriteLine("  srcQuery      - Source query");
      Console.WriteLine("  dstServer     - Destination database connection string");
      Console.WriteLine("  dstTable      - Desitination table");
       Console.WriteLine(" truncate      - Optional, to truncate destination table");
      Console.WriteLine();
      Console.WriteLine("Example> Gen_SSIS_Pkg_Query2Table \"Data Source=EPICCLSQLPRD;Initial Catalog=Clarity;Provider=SQLNCLI11.1;Integrated Security=SSPI;Auto Translate=False;\" \"select department_id as id, department_name as name from clarity_dep\" \"Data Source=EPICCDWDBDEV;Initial Catalog=JMHStage;Provider=SQLNCLI11.1;Integrated Security=SSPI;Auto Translate=False;\" ");
      Console.WriteLine();
      Console.WriteLine("The generated package can then run to load the data.");
      Console.WriteLine("The column names in the source system query must match the destination table column names, must column alias if different.");
      //Console.WriteLine("If a full refresh is needed, you'll have to truncate the table first.");
      Console.WriteLine();
      Console.WriteLine("Example> dtexec /f Load_TableA.dtsx");
      return;
    }
    var srcConnStr  = args[0];
    var srcQuery    = args[1];
    var dstConnStr  = args[2];
    var dstTable    = args[3];
    var truncateFlg = "NA";
    
    if (args.Length == 5)
    {
      truncateFlg = args[4];
    }

    Console.WriteLine("Building Package...");

    //Creating the package
    Package pkg = new Package();
    
    // Create and assign Variables
    Variable varSrcConnStr  = pkg.Variables.Add("srcConnectStr", false, "User", srcConnStr);
    Variable varSrcQuery    = pkg.Variables.Add("srcQuery", false, "User", srcQuery);
    Variable varDstConnStr  = pkg.Variables.Add("dstConnectStr", false, "User", dstConnStr);
    Variable varDstTable    = pkg.Variables.Add("dstTable", false, "User", dstTable);
    

    Executable sqlTaskExe = null;
    TaskHost hostSqlTask = null;
    if (truncateFlg == "truncate")
    {
      Console.WriteLine("Adding Execute SQL Task");
      sqlTaskExe = pkg.Executables.Add("STOCK:SQLTask");
      hostSqlTask = (TaskHost)sqlTaskExe;
      hostSqlTask.Properties["Name"].SetValue(hostSqlTask, "Execute SQL Task - truncate table " + dstTable);
      hostSqlTask.Properties["Connection"].SetValue(hostSqlTask, "OLEDB Destination ConnectionManager");
      hostSqlTask.Properties["SqlStatementSource"].SetValue(hostSqlTask, "truncate table " + dstTable);
    }
                         
    Console.WriteLine("Adding Data Flow Task");
    Executable exe = pkg.Executables.Add("STOCK:PipelineTask");
    TaskHost hostDft = (TaskHost)exe;
    hostDft.Name = "DFT - Transfer Departments";
    MainPipe dft = (hostDft).InnerObject as MainPipe;
    
    if (truncateFlg == "truncate")
    {
      Console.WriteLine("Connect ExecuteSqlTask to DFT");
      PrecedenceConstraint pcFileTasks =  pkg.PrecedenceConstraints.Add((Executable)hostSqlTask, (Executable)hostDft);
      pcFileTasks.Value = DTSExecResult.Completion;
    }

    Console.WriteLine("Creating Source OLE DB Connection Manager");
    ConnectionManager cnSource = pkg.Connections.Add("OLEDB");
    cnSource.Name = "OLE DB Source ConnectionManager";
    cnSource.SetExpression("ConnectionString", "@[User::srcConnectStr]");
    cnSource.DelayValidation = true;
    //cnSource.ConnectionString = string.Format("Provider=SQLOLEDB.1;Data Source={0};Initial Catalog={1};Integrated Security=SSPI;", varSrcServer.Value.ToString(), varSrcDatabase.Value.ToString());
    //cnSource.SetExpression("ServerName", "@[User::srcServer]");
    //cnSource.SetExpression("InitialCatalog", "@[User::srcDatabase]");

    Console.WriteLine("Creating Destination OLE DB Connection Manager");
    ConnectionManager cnDestination = pkg.Connections.Add("OLEDB");
    cnDestination.Name = "OLEDB Destination ConnectionManager";
    //cnDestination.ConnectionString = string.Format("Provider=SQLOLEDB.1;Data Source={0};Initial Catalog={1};Integrated Security=SSPI;", varDstServer.Value.ToString(), varDstDatabase.Value.ToString());
    cnDestination.SetExpression("ConnectionString", "@[User::dstConnectStr]");

    Console.WriteLine("Adding OLE DB Source");
    IDTSComponentMetaData100 component =
    dft.ComponentMetaDataCollection.New();
    component.Name = "OLEDBSource";
    component.ComponentClassID = "DTSAdapter.OleDbSource.3";

    Console.WriteLine("Initialize");
    CManagedComponentWrapper instance = component.Instantiate();
    instance.ProvideComponentProperties();

    Console.WriteLine("Assign Connection Manager to Component");
    component.RuntimeConnectionCollection[0].ConnectionManager = DtsConvert.GetExtendedInterface(pkg.Connections[0]);
    component.RuntimeConnectionCollection[0].ConnectionManagerID = pkg.Connections[0].ID;

    Console.WriteLine("Set other properties");
    //instance.SetComponentProperty("AccessMode", 2);
    //instance.SetComponentProperty("SqlCommand", varSrcQuery.Value.ToString());
    instance.SetComponentProperty("AccessMode", 3);
    instance.SetComponentProperty("SqlCommandVariable", "User::srcQuery");

    Console.WriteLine("Set other properties");
    instance.AcquireConnections(null);
    instance.ReinitializeMetaData();
    instance.ReleaseConnections();

    Console.WriteLine("Add OLE DB Destination");
    IDTSComponentMetaData100 destination = dft.ComponentMetaDataCollection.New();
    destination.ComponentClassID = "DTSAdapter.OleDbDestination";
    destination.Name = "OLEDBDestination";

    Console.WriteLine("Instantiate Destination Component");
    CManagedComponentWrapper destInstance = destination.Instantiate();
    destInstance.ProvideComponentProperties();

    Console.WriteLine("Assign Destination Connection");
    destination.RuntimeConnectionCollection[0].ConnectionManager =
    DtsConvert.GetExtendedInterface(pkg.Connections[1]);
    destination.RuntimeConnectionCollection[0].ConnectionManagerID = pkg.Connections[1].ID;

    Console.WriteLine("Set destination properties");
    //destInstance.SetComponentProperty("OpenRowset", varDstTable.Value.ToString());
    //destInstance.SetComponentProperty("AccessMode", 3);
    destInstance.SetComponentProperty("OpenRowsetVariable", "User::dstTable");
    destInstance.SetComponentProperty("AccessMode", 4);
    destInstance.SetComponentProperty("FastLoadOptions", "TABLOCK,CHECK_CONSTRAINTS");

    Console.WriteLine("Connect source to destination");
    dft.PathCollection.New().AttachPathAndPropagateNotifications(component.OutputCollection[0], destination.InputCollection[0]);

    Console.WriteLine("Reinitialize destination");
    destInstance.AcquireConnections(null);
    destInstance.ReinitializeMetaData();
    destInstance.ReleaseConnections();

    Console.WriteLine("Map input columns with virtual input columns");
    IDTSInput100 input = destination.InputCollection[0];
    IDTSVirtualInput100 vInput = input.GetVirtualInput();

    foreach (IDTSVirtualInputColumn100 vColumn in vInput.VirtualInputColumnCollection)
    {
        // Select column, and retain new input column
        IDTSInputColumn100 inputColumn = destInstance.SetUsageType(input.ID,
        vInput, vColumn.LineageID, DTSUsageType.UT_READONLY);
        // Find external column by name
        IDTSExternalMetadataColumn100 externalColumn = input.ExternalMetadataColumnCollection[inputColumn.Name];
        // Map input column to external column
        destInstance.MapInputColumn(input.ID, inputColumn.ID, externalColumn.ID);
    }
    
    Console.WriteLine("Finished mapping input columns with virtual input columns");
            
    //Execute Package
    // DTSExecResult result = pkg.Execute();
    // Console.WriteLine(result.ToString());
    
    var dtsx = new StringBuilder();
    dtsx.Append(Directory.GetCurrentDirectory()).Append("\\").Append("Load_Query2").Append(dstTable).Append(".dtsx");

    var app = new Application();
    Console.WriteLine("Saving Package {0}", dtsx.ToString());
    app.SaveToXml(dtsx.ToString(), pkg, null);
       
    pkg.Dispose();
    Console.WriteLine("Done");
          
  } // Main()
        
  } // class LoadTableFromQuery
    
} // namespace JMH_SSIS_Gen_Query2Table
