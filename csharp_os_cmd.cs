/*
 * csharp_os_cmd.cs
 * Craig Nobili
 */

using System;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.IO;

public class StoredProcedures
{

//System.Diagnostics.Process.Start(@"C:\MyScript.bat");

public static void RunCmdSync(object command)
{
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

public static void RunCmdAsync(string command)
{
  //Asynchronously start the Thread to process the Execute command request.
  Thread objThread = new Thread(new ParameterizedThreadStart(RunCmdSync));
  //Make the thread as background thread.
  objThread.IsBackground = true;
  //Set the Priority of the thread.
  objThread.Priority = ThreadPriority.AboveNormal;
  //Start the thread.
  objThread.Start(command);

} // RunCmdAsync()

public static void Main()
{
  //RunCmdSync(@"c:\C#\DumpData\DumpData2 c:\C#\DumpData\POC.cfg");
  RunCmdAsync(@"c:\C#\DumpData\DumpData2 c:\C#\DumpData\POC.cfg");

} // Main()

} // class StoredProcedures