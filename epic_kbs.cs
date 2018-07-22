using System;
using System.Data;
using System.Data.Common;
using KBSql.Data.KBSqlClient;

namespace KBSqlEx
{
	public class KBSqlEx
	{
		public static void Main(string[] args)
		{
			// Setup connection string to access Oracle database

			string connectString="host=epic_nonprd1; port=6006; applicationname=TestApp; userid=user; password=passwd";

			try
			{
				// Instantiate the connection, passing the connection string into the constructor
				KBSDbConnection con = new KBSDbConnection(connectString);

				// Open the connection
				con.Open();

				// Create and execute and execute the query
				DbCommand cmd = con.CreateCommand();
				cmd.CommandText ="select pat_name from ept.noadd_single";
				DbDataReader reader = cmd.ExecuteReader();

				// Iterate through the DataReader and display row
				while (reader.Read())
				{
					Console.WriteLine("{0}", reader.GetString(0));
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("Here");
				Console.WriteLine(e.ToString());
			}

		} // main()
	} // Class KBSqlEx
} // namespace KBSqlEx
