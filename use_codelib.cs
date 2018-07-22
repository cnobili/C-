/*
 * use_codelib.cs
 *
 * Craig Nobili
 */

using System;
using System.IO;
using System.Text;
using CodeLibrary;

public class use_log
{
  public static void Main(String [] args)
  {
    //Util.LogInit(null, "test", Util.LogLevel.DEBUG);

    //Util.LogMsg(Util.LogLevel.DEBUG, "testing DEBUG");
    //Util.LogMsg(Util.LogLevel.NORM, "testing NORM");
    //Util.LogMsg(Util.LogLevel.MUST, "testing MUST");

    //Util.LogClose();

    //String[] addresses = new String[] {"craig.nobili@johnmuirhealth.com"};
    //Util.SendMail("ClarityMonitor@johnmuirhealth.com", addresses, "test", "<a href='http://www.abc.com' target='_blank'>www.abc.com</a>");
    
    //String srcConnStr = Util.GetSqlConnStr("epicclsqlprd", "Clarity", "clarityreport", "A&n2kw(");
    //String dstConnStr = Util.GetSqlConnStr("revdashdbdev", "ClarityMonitor", "clmon", "clmon4jmh");
    //Util.BulkCopy(srcConnStr, "select * from cr_stat_err_codes", dstConnStr, "oat", 60);
    
    String srcConnStr = "Data Source=EPICCDWDBDEV;Initial Catalog=EDW;Integrated Security=SSPI;";
    String dstConnStr = "Data Source=EPICCDWDBDEV;Initial Catalog=JMHStage;Integrated Security=SSPI;";
    Util.BulkCopy(srcConnStr, "select * from ProviderDim", dstConnStr, "ProviderDim", 60 * 5);
    
    //Console.Write("Enter password:");
    //String str = Util.ReadPassword();
    //Console.WriteLine("\nYou entered <{0}>", str);
    
    //Util.GenTsqlTable("c:\\c#\\CodeLibrary\\orders2.dat", "|", "mytable");
    //Util.GenTsqlBulkIns("c:\\c#\\CodeLibrary\\orders2.dat", "mytable", "|");
    //Util.GenTsqlReadFileAsTable("c:\\work\\JMH\\c#\\CodeLibrary\\orders2.dat", "|");
    //Util.GenTsqlReadFileAsTable("c:\\c#\\CodeLibrary\\orders2.dat", "|", "\\r\\n");
    
    /*
    Util.GenOraExtTab
    (
      "c:\\c#\\CodeLibrary\\orders2.dat"
    , "|"
    , "mytable_ext"
    , "ORA_DIR_OBJ"
    , "N"
    );
    */

  } // Main()

} // use_log class
