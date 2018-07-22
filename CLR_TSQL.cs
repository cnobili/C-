/*
 * CLR_TSQL.cs
 *
 * Stored Procedures and Functions for CLR Integration with SQL Server.
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
using System.Text.RegularExpressions;

public class CLRTSQL
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
  
  [Microsoft.SqlServer.Server.SqlFunction]
  public static int IsMatch(String str, String regex) 
  {
    Match m;
    
    try
    {
      m = Regex.Match(str, regex, RegexOptions.IgnoreCase | RegexOptions.Compiled);
            
      if (m.Success)
        return(1);
      else
        return(0);
    }
    catch (Exception e)
    {
      SqlContext.Pipe.Send(e.ToString());
      return(0);
    }
                   
  } // IsMatch()
  
  [Microsoft.SqlServer.Server.SqlProcedure]
  public static void SqlExtractData(String outputfile, String delim, String columnHeader, String query, String server, String database, String user , String pass)
  {
    String dbConnectStr;
    StreamWriter pw = null;
    SqlConnection dbConn = null;
        
    if (user == null)
    {
      dbConnectStr = "Server=" + server + ";Database=" + database + ";Trusted_Connection=True";
    }
    else
    {
      dbConnectStr = "Server=" + server + ";Database=" + database + ";user id=" + user + ";password=" + pass;
    }
       
    try
    {
      dbConn = new SqlConnection(dbConnectStr);
      dbConn.Open();
      Console.WriteLine("Opened database connection");
      int recs = 0;
      SqlCommand cmdSQL = new SqlCommand(query, dbConn);
      cmdSQL.CommandTimeout = 0;
      SqlDataReader dataReader = cmdSQL.ExecuteReader();
      int fieldCount = dataReader.FieldCount;
 
      String rec = null;
      String fieldSep = null;
      String col = null;
      int displayCount = 100000;

      Console.WriteLine("Execute SQL = " + query);
      Console.WriteLine("Number of columns in select stmt = " + fieldCount);

      pw = new StreamWriter(outputfile);

      if (columnHeader.Equals("Y"))
      {
        rec = "";
        fieldSep = "";
        for (int i = 0; i < fieldCount; i++)
        {
          rec += fieldSep;
          rec += dataReader.GetName(i);
          fieldSep = delim;
        }
        pw.WriteLine(rec);
      }

      // Iterate through the resultset
      while (dataReader.Read())
      {
        rec = "";
        fieldSep = "";
        for (int i = 0; i < fieldCount; i++)
        {
          rec += fieldSep;
          col = dataReader[i].ToString();
          if (col != null)
            rec += col;
          else
            rec += "";
          fieldSep = delim;
        }

        recs++;

        if (recs % displayCount == 0)
          Console.WriteLine("Total records written so far = " + recs);
        pw.WriteLine(rec);
      }

      Console.WriteLine("\nTotal records written out to file = " + recs);

    }
    catch(Exception e)
    {
      Console.WriteLine(e.ToString());
    }
    finally
    {
      pw.Close();
      dbConn.Close();
    }

  } // SqlExtractData()

} // class CLRTSQL
