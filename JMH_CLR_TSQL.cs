/*
 * JMH_CLR_TSQL.cs
 *
 * Stored Procedures for CLR Integration with SQL Server.
 *
 * Craig Nobili
 */

using System;
//using System.Threading;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;

public class StoredProcedures
{

  [Microsoft.SqlServer.Server.SqlProcedure]
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

} // class StoredProcedures
