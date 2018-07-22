/*
 * CodeLib.cs
 *
 * Library of common methods.
 *
 * Craig Nobili
 */

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace CodeLibrary
{

public class Util
{

  // Private Members
  public const String MAIL_SERVER = "mr1.hsys.local";
  
  static readonly string PasswordHash = "P@@Sw0rd";
  static readonly string SaltKey = "S@LT&KEY";
  static readonly string VIKey = "@1B2c3D4e5F6g7H8";

  // Enums
  public enum LogLevel {DEBUG = 0, NORM, MUST}

  private static LogLevel sysLevel;
  private static StreamWriter pw = null;
  
  public const int MAX_SQL_COL_NAME_LEN = 60;
  public const int MAX_ORA_COL_NAME_LEN = 30;
  
  public const String SQL_SUFFIX           = ".sql";
  public const String MSQL_BULK_INS_PREFIX = "bulkIns_";
  public const String MSQL_FILE_SEL_PREFIX = "fileSel_";
  
  public const String BCP_VERSION          = "9.0"; // SQL 2005
  public const String TAB                  = "\t";
  public const String FILE_DATA_TYPE       = "SQLCHAR";

  public static void LogInit(String logDir, String logPrefix, LogLevel level )
  {
    String logfile;
    sysLevel = level;

    if (logPrefix != null)
    {
      logfile = logPrefix + "_" + DateTime.Now.ToString("yyyyMMdd") + ".log";

      if (logDir != null)
      {
        if ( logDir[logDir.Length - 1] == '\\' )
          logfile = logDir + logfile;
        else
          logfile = logDir + "\\" + logfile;
      }
      pw = new StreamWriter(logfile);
    }

  } // LogInit()

  public static void LogClose()
  {
    if (pw != null)
      pw.Close();

  } // LogClose()

  public static void LogMsg(LogLevel level, String s)
  {
    // Only log message where level >= sysLevel
    if (level < sysLevel) return;

    Console.WriteLine("{0} => {1}", DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss"), s);
    if (pw != null)
    {
      pw.WriteLine("{0} => {1}", DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss"), s);
      pw.Flush(); // so output is written to file immediately
    }

  } // LogMsg()

  public static void SendMail(String from, String[] to, String subject, String body)
  {
    MailMessage msg = new MailMessage();
    SmtpClient SmtpServer = new SmtpClient(MAIL_SERVER);

    msg.From = new MailAddress(from);
    foreach (String s in to)
    {
      msg.To.Add(s);
    }

    msg.Subject = subject;
    msg.IsBodyHtml = true;
    msg.Body = body;

    SmtpServer.Port = 25;
    SmtpServer.Send(msg);

  } // SendMail()

  /*
   * GetSqlConnStr()
   *
   * Returns SQL Serve Connection String.
   */
  public static String GetSqlConnStr(String server, String database, String userid, String passwd) 
  {
    String sqlConnStr = "user id=" + userid + "; password=" + passwd 
      + "; server=" + server + "; Trusted_Connection=false; database=" + database;
        
    return(sqlConnStr);
  
  } // GetSqlConnStr()
  
  /*
   * BulkCopy() - For SQL Server, source and destination tables need to be the same exact structure.
   */
  public static void BulkCopy
  (
    String srcConnStr
  , String srcQuery
  , String dstConnStr
  , String dstTableName
  , int    timeoutInSecs         
  )
  {
    // Open a sourceConnection
    using (SqlConnection sourceConnection = new SqlConnection(srcConnStr))
    {
      sourceConnection.Open();

      // Get data from the source table as a SqlDataReader.
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
        
        // Connection will be automatically closed upon exit of using block
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
        {
          bulkCopy.DestinationTableName = dstTableName;
          bulkCopy.BulkCopyTimeout = timeoutInSecs; // if not specified defaults to 30 secs

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
      
      }
    }   
      
  } // BulkCopy()
  
  public static string Encrypt(string plainText)
  {
    byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

    byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
    var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
    var encryptor = symmetricKey.CreateEncryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));

    byte[] cipherTextBytes;

    using (var memoryStream = new MemoryStream())
    {
      using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
      {
        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
        cryptoStream.FlushFinalBlock();
        cipherTextBytes = memoryStream.ToArray();
        cryptoStream.Close();
      }
      memoryStream.Close();
    }
    return Convert.ToBase64String(cipherTextBytes);

  } // Encrypt()

  public static string Decrypt(string encryptedText)
  {
    byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
    byte[] keyBytes = new Rfc2898DeriveBytes(PasswordHash, Encoding.ASCII.GetBytes(SaltKey)).GetBytes(256 / 8);
    var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

    var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(VIKey));
    var memoryStream = new MemoryStream(cipherTextBytes);
    var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
    byte[] plainTextBytes = new byte[cipherTextBytes.Length];

    int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
    memoryStream.Close();
    cryptoStream.Close();
    return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());

  } // Decrypt()

  // Mask password as it is read in from command line prompt
  public static string ReadPassword()
  {
    Stack<string> pass = new Stack<string>();

    for (ConsoleKeyInfo consKeyInfo = Console.ReadKey(true); consKeyInfo.Key != ConsoleKey.Enter; consKeyInfo = Console.ReadKey(true))
    {
      if (consKeyInfo.Key == ConsoleKey.Backspace)
      {
        try
        {
          Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
          Console.Write(" ");
          Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
          pass.Pop();
        }
        catch (InvalidOperationException ex)
        {
          /* Nothing to delete, go back to previous position */
          Console.SetCursorPosition(Console.CursorLeft + 1, Console.CursorTop);
        }
      }
      else
      {
        Console.Write("*");
        pass.Push(consKeyInfo.KeyChar.ToString());
      }
    }

    String[] password = pass.ToArray();
    Array.Reverse(password);
    return string.Join(string.Empty, password);

  } // ReadPassword()
  
  public static String PadRight(String s, int n)
  {
    return String.Format("{0," + -n + "}", s);

  } // PadRight()

  public static void GenTsqlTable(String dataFile, String delim, String msqlTableName)
  {
    String colDef;
    String fileDir = Path.GetDirectoryName(dataFile);
    FileInfo fi = new FileInfo(dataFile);
    StreamReader sr = fi.OpenText();
    String header = sr.ReadLine();
    sr.Close();

    String [] tokens = header.Split(new Char [] {System.Convert.ToChar(delim)});

    StreamWriter pw = new StreamWriter(fileDir + "\\" + msqlTableName + SQL_SUFFIX);

    pw.WriteLine("create table " + msqlTableName);
    pw.WriteLine("(");

    for (int i = 0; i < tokens.Length; i++)
    {
      colDef = tokens[i].ToLower();
      if (colDef.Length > MAX_SQL_COL_NAME_LEN)
      {
        colDef.Substring(0, MAX_SQL_COL_NAME_LEN);
      }

      colDef = PadRight(colDef, MAX_SQL_COL_NAME_LEN) + " varchar(8000)";

      if (i == 0)
      {
        pw.WriteLine("  " + colDef);
      }
      else
      {
        pw.WriteLine(", " + colDef);
      }
    }

    pw.WriteLine(")");

    pw.Close();

  } // GenTsqlTable()
  
  //public static void GenTsqlFormatFile(String dataFile, String colDelim, String rowDelim = "\\r\\n")
  public static void GenTsqlFormatFile(String dataFile, String colDelim, String rowDelim)
  {
    String fileDir = Path.GetDirectoryName(dataFile);
    String fileNameNoExt = Path.GetFileNameWithoutExtension(dataFile);
    FileInfo fi = new FileInfo(dataFile);
    StreamReader sr = fi.OpenText();
    String header = sr.ReadLine();
    sr.Close();

    String [] tokens = header.Split(new Char [] {System.Convert.ToChar(colDelim)});

    StreamWriter pw = new StreamWriter(fileDir + "\\" + fileNameNoExt + ".fmt");
    
    pw.WriteLine(BCP_VERSION);
    pw.WriteLine(tokens.Length); // number of columns
    
    for (int i = 0; i < tokens.Length; i++)
    {
      pw.Write("{0}{1}", i + 1, TAB);          // field order in host file
      pw.Write("{0}{1}", FILE_DATA_TYPE, TAB); // host file data type
      pw.Write("{0}{1}", 0, TAB);              // prefix length
      pw.Write("{0}{1}", 0, TAB);              // host file data length (use zero)
      if (i == tokens.Length - 1)
        pw.Write("{0}{1}", rowDelim, TAB);     // last column delimiter (i.e. row delimiter)
      else
        pw.Write("{0}{1}", colDelim, TAB);     // column delimiter
      pw.Write("{0}{1}", i + 1, TAB);          // server column order
      pw.Write("{0}{1}", tokens[i], TAB);      // server column name
      pw.Write("{0}", "\"\"");                 // column collation
      pw.WriteLine();
    }
    
    pw.Close();
  
  } // GenTsqlFormatFile()

  public static void GenTsqlBulkIns(String dataFile, String msqlTableName, String delim)
  {
    String outputDir = Path.GetDirectoryName(dataFile);
    StreamWriter pw = new StreamWriter(outputDir + "\\" + MSQL_BULK_INS_PREFIX + msqlTableName + SQL_SUFFIX);

    pw.WriteLine("bulk insert " + msqlTableName);
    pw.WriteLine("from '" + dataFile + "'");
    pw.WriteLine("with");
    pw.WriteLine("(");
    pw.WriteLine("  datafiletype = 'char'");
    pw.WriteLine(", firstrow = 2");
    pw.WriteLine(", fieldterminator = '" + delim + "'");
    pw.WriteLine(", errorfile = '" + outputDir + "\\\\" + msqlTableName + ".err'");
    pw.WriteLine(")");

    pw.Close();

  } // GenTsqlBulkIns()
  
  //public static void GenTsqlReadFileAsTable(String dataFile, String colDelim, String rowDelim = "\\r\\n")
  public static void GenTsqlReadFileAsTable(String dataFile, String colDelim, String rowDelim)
  {
    String outputDir = Path.GetDirectoryName(dataFile);
    String filename = Path.GetFileName(dataFile);
    String filenameNoExt = Path.GetFileNameWithoutExtension(dataFile);
    
    StreamWriter pw = new StreamWriter(outputDir + "\\" + MSQL_FILE_SEL_PREFIX + filenameNoExt + SQL_SUFFIX);
    
    pw.WriteLine("select *");
    pw.WriteLine("from openrowset");
    pw.WriteLine("(");
    pw.WriteLine("  BULK '{0}'", dataFile);
    pw.WriteLine(", formatfile = '{0}\\{1}.fmt'", outputDir, filenameNoExt);  
    pw.WriteLine(", firstrow = 2");
    pw.WriteLine(", errorfile = '{0}\\{1}.err'", outputDir, filenameNoExt);
    pw.WriteLine(") as {0}", filenameNoExt);
    
    pw.Close();
    
    // Generate format file
    GenTsqlFormatFile(dataFile, colDelim, rowDelim);
      
  } // GenTsqlReadFileAsTable()
  
  public static void GenOraExtTab
  (
    String dataFile
  , String colDelim
  , String oraExtTableName
  , String oraExtDirectoryName
  , String oraExtAddRecNum
  )
  {
    String filename = Path.GetFileName(dataFile);
    String outputDir = Path.GetDirectoryName(dataFile);
    
    String colDef;  
    FileInfo fi = new FileInfo(dataFile);
    StreamReader sr = fi.OpenText();
    String header = sr.ReadLine();
    sr.Close();

    String [] tokens = header.Split(new Char [] {System.Convert.ToChar(colDelim)});
  
    StreamWriter pw = new StreamWriter(outputDir + "\\" + oraExtTableName + SQL_SUFFIX);
    
    pw.WriteLine("create table " + oraExtTableName);
    pw.WriteLine("(");

    for (int i = 0; i < tokens.Length; i++)
    {
      if (tokens[i].Length == 0) continue;

      colDef = tokens[i].ToLower();
      colDef = Regex.Replace(colDef, @"[- ,']", "_");

      if (colDef.Length > MAX_ORA_COL_NAME_LEN)
      {
        colDef = colDef.Substring(0, MAX_ORA_COL_NAME_LEN);
      }

      colDef = PadRight(colDef, MAX_ORA_COL_NAME_LEN) + " varchar(4000)";

      if (i == 0)
      {
        pw.WriteLine("  " + colDef);
      }
      else
      {
        pw.WriteLine(", " + colDef);
      }
    }

    if (oraExtAddRecNum.Equals("Y"))
    {
      colDef = PadRight("rec_num", MAX_ORA_COL_NAME_LEN) + " number";
      pw.WriteLine(", " + colDef);
    }

    pw.WriteLine(")");
    pw.WriteLine("organization external");
    pw.WriteLine("(");
    pw.WriteLine("  type oracle_loader");
    pw.WriteLine("  default directory " + oraExtDirectoryName.ToUpper());
    pw.WriteLine("  access parameters");
    pw.WriteLine("  (");
    pw.WriteLine("    records delimited by NEWLINE");
    pw.WriteLine("    readsize 5000000");
    pw.WriteLine("    skip 1");
    pw.WriteLine("    nobadfile");
    pw.WriteLine("    nodiscardfile");
    pw.WriteLine("    nologfile");
    pw.WriteLine("    fields terminated by '" + colDelim + "'");
    pw.WriteLine("    missing field values are null");
    pw.WriteLine("    reject rows with all null fields");

    if (oraExtAddRecNum .Equals("Y"))
    {
      pw.WriteLine("    (");

      for (int i = 0; i < tokens.Length; i++)
      {
        if (tokens[i].Length == 0) continue;

        colDef = tokens[i].ToLower();
        Regex.Replace(colDef, @"[- ,']", "_");

        if (colDef.Length > MAX_ORA_COL_NAME_LEN)
        {
          colDef = colDef.Substring(0, MAX_ORA_COL_NAME_LEN);
        }

        colDef = PadRight(colDef, MAX_ORA_COL_NAME_LEN);

        if (i == 0)
        {
          pw.WriteLine("      " + colDef);
        }
        else
        {
          pw.WriteLine("    , " + colDef);
        }
      }

      colDef = PadRight("rec_num", MAX_ORA_COL_NAME_LEN) + " recnum";
      pw.WriteLine("    , " + colDef);
      pw.WriteLine("    )");
    }

    pw.WriteLine("  )");
    pw.WriteLine("  location ('" + filename + "')");
    pw.WriteLine(")");
    pw.WriteLine("reject limit unlimited");
    pw.WriteLine("parallel");
    pw.WriteLine(";");

    pw.Close();

  } // GenOraExtTab()

} // class Util

} // namespace CodeLibrary