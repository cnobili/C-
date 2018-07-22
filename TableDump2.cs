using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OracleClient;
//using Oracle.DataAccess.Client;
//using Oracle.DataAccess.Types;
using System.IO;

namespace TableDump2
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class TableDump : System.Windows.Forms.Form
	{
    private System.Windows.Forms.Label tnsLabel;
    private System.Windows.Forms.Label userLabel;
    private System.Windows.Forms.Label passwdLabel;
    private System.Windows.Forms.Label schemaLabel;
    private System.Windows.Forms.Label tableLabel;
    private System.Windows.Forms.Label whereClauseLabel;
    private System.Windows.Forms.Label DelimiterLabel;
    private System.Windows.Forms.TextBox tnsTextBox;
    private System.Windows.Forms.TextBox userTextBox;
    private System.Windows.Forms.TextBox passwdTextBox;
    private System.Windows.Forms.TextBox schemaTextBox;
    private System.Windows.Forms.TextBox tableTextBox;
    private System.Windows.Forms.TextBox whereClauseTextBox;
    private System.Windows.Forms.TextBox delimiterTextBox;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

    private const string DATE     = "Date";
    private string tnsService;
    private string userid;
    private string passwd;
    private string schema;
    private string table;
    private string whereClause;
    private System.Windows.Forms.Button submitButton;
    private System.Windows.Forms.Label outputLabel;
    private string delimiter;
    private OracleConnection oraCon;

		public TableDump()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
      this.tnsLabel = new System.Windows.Forms.Label();
      this.userLabel = new System.Windows.Forms.Label();
      this.passwdLabel = new System.Windows.Forms.Label();
      this.schemaLabel = new System.Windows.Forms.Label();
      this.tableLabel = new System.Windows.Forms.Label();
      this.whereClauseLabel = new System.Windows.Forms.Label();
      this.DelimiterLabel = new System.Windows.Forms.Label();
      this.tnsTextBox = new System.Windows.Forms.TextBox();
      this.userTextBox = new System.Windows.Forms.TextBox();
      this.passwdTextBox = new System.Windows.Forms.TextBox();
      this.schemaTextBox = new System.Windows.Forms.TextBox();
      this.tableTextBox = new System.Windows.Forms.TextBox();
      this.whereClauseTextBox = new System.Windows.Forms.TextBox();
      this.delimiterTextBox = new System.Windows.Forms.TextBox();
      this.submitButton = new System.Windows.Forms.Button();
      this.outputLabel = new System.Windows.Forms.Label();
      this.SuspendLayout();
      //
      // tnsLabel
      //
      this.tnsLabel.Location = new System.Drawing.Point(16, 40);
      this.tnsLabel.Name = "tnsLabel";
      this.tnsLabel.Size = new System.Drawing.Size(112, 23);
      this.tnsLabel.TabIndex = 0;
      this.tnsLabel.Text = "TNS Service";
      //
      // userLabel
      //
      this.userLabel.Location = new System.Drawing.Point(16, 72);
      this.userLabel.Name = "userLabel";
      this.userLabel.TabIndex = 1;
      this.userLabel.Text = "UserID";
      //
      // passwdLabel
      //
      this.passwdLabel.Location = new System.Drawing.Point(16, 104);
      this.passwdLabel.Name = "passwdLabel";
      this.passwdLabel.TabIndex = 2;
      this.passwdLabel.Text = "Password";
      //
      // schemaLabel
      //
      this.schemaLabel.Location = new System.Drawing.Point(16, 136);
      this.schemaLabel.Name = "schemaLabel";
      this.schemaLabel.TabIndex = 3;
      this.schemaLabel.Text = "Schema";
      //
      // tableLabel
      //
      this.tableLabel.Location = new System.Drawing.Point(16, 168);
      this.tableLabel.Name = "tableLabel";
      this.tableLabel.TabIndex = 4;
      this.tableLabel.Text = "Table";
      //
      // whereClauseLabel
      //
      this.whereClauseLabel.Location = new System.Drawing.Point(16, 200);
      this.whereClauseLabel.Name = "whereClauseLabel";
      this.whereClauseLabel.TabIndex = 5;
      this.whereClauseLabel.Text = "Where Clause";
      //
      // DelimiterLabel
      //
      this.DelimiterLabel.Location = new System.Drawing.Point(16, 232);
      this.DelimiterLabel.Name = "DelimiterLabel";
      this.DelimiterLabel.TabIndex = 6;
      this.DelimiterLabel.Text = "Delimiter";
      //
      // tnsTextBox
      //
      this.tnsTextBox.Location = new System.Drawing.Point(120, 40);
      this.tnsTextBox.Name = "tnsTextBox";
      this.tnsTextBox.TabIndex = 7;
      this.tnsTextBox.Text = "";
      //
      // userTextBox
      //
      this.userTextBox.Location = new System.Drawing.Point(120, 72);
      this.userTextBox.Name = "userTextBox";
      this.userTextBox.TabIndex = 8;
      this.userTextBox.Text = "";
      //
      // passwdTextBox
      //
      this.passwdTextBox.Location = new System.Drawing.Point(120, 104);
      this.passwdTextBox.Name = "passwdTextBox";
      this.passwdTextBox.PasswordChar = '*';
      this.passwdTextBox.TabIndex = 9;
      this.passwdTextBox.Text = "";
      //
      // schemaTextBox
      //
      this.schemaTextBox.Location = new System.Drawing.Point(120, 136);
      this.schemaTextBox.Name = "schemaTextBox";
      this.schemaTextBox.TabIndex = 10;
      this.schemaTextBox.Text = "";
      //
      // tableTextBox
      //
      this.tableTextBox.Location = new System.Drawing.Point(120, 168);
      this.tableTextBox.Name = "tableTextBox";
      this.tableTextBox.TabIndex = 11;
      this.tableTextBox.Text = "";
      //
      // whereClauseTextBox
      //
      this.whereClauseTextBox.Location = new System.Drawing.Point(120, 200);
      this.whereClauseTextBox.Name = "whereClauseTextBox";
      this.whereClauseTextBox.TabIndex = 12;
      this.whereClauseTextBox.Text = "";
      //
      // delimiterTextBox
      //
      this.delimiterTextBox.Location = new System.Drawing.Point(120, 232);
      this.delimiterTextBox.Name = "delimiterTextBox";
      this.delimiterTextBox.TabIndex = 13;
      this.delimiterTextBox.Text = "";
      //
      // submitButton
      //
      this.submitButton.Location = new System.Drawing.Point(24, 312);
      this.submitButton.Name = "submitButton";
      this.submitButton.TabIndex = 14;
      this.submitButton.Text = "Submit";
      this.submitButton.Click += new System.EventHandler(this.submitButton_Click);
      //
      // outputLabel
      //
      this.outputLabel.Location = new System.Drawing.Point(328, 312);
      this.outputLabel.Name = "outputLabel";
      this.outputLabel.Size = new System.Drawing.Size(224, 40);
      this.outputLabel.TabIndex = 15;
      //
      // TableDump
      //
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.ClientSize = new System.Drawing.Size(624, 413);
      this.Controls.Add(this.outputLabel);
      this.Controls.Add(this.submitButton);
      this.Controls.Add(this.delimiterTextBox);
      this.Controls.Add(this.whereClauseTextBox);
      this.Controls.Add(this.tableTextBox);
      this.Controls.Add(this.schemaTextBox);
      this.Controls.Add(this.passwdTextBox);
      this.Controls.Add(this.userTextBox);
      this.Controls.Add(this.tnsTextBox);
      this.Controls.Add(this.DelimiterLabel);
      this.Controls.Add(this.whereClauseLabel);
      this.Controls.Add(this.tableLabel);
      this.Controls.Add(this.schemaLabel);
      this.Controls.Add(this.passwdLabel);
      this.Controls.Add(this.userLabel);
      this.Controls.Add(this.tnsLabel);
      this.Name = "TableDump";
      this.Text = "TableDump";
      this.Load += new System.EventHandler(this.TableDump_Load);
      this.ResumeLayout(false);

    }
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.Run(new TableDump());
		}

    private void TableDump_Load(object sender, System.EventArgs e)
    {
    }

    private void submitButton_Click(object sender, System.EventArgs e)
    {
      tnsService = tnsTextBox.Text;
      userid = userTextBox.Text;
      passwd = passwdTextBox.Text;
      schema = schemaTextBox.Text;
      table = tableTextBox.Text;
      whereClause = whereClauseTextBox.Text;
      delimiter = delimiterTextBox.Text;

      // Connect to Oracle and get the data
      connectToDB();
      dumpData();
      oraCon.Close();
      oraCon.Dispose();
    }

    private void connectToDB()
    {
      string connStr = "User Id=" + userid + ";";
      connStr += "Password=" + passwd + ";";
      connStr += "Data Source=" + tnsService;

      // Connect to Oracle
      oraCon = new OracleConnection(connStr);
      try
      {
        oraCon.Open();
      }
      catch (Exception ex)
      {
        outputLabel.Text = "Failed to Connect to Oracle:" + ex.Message;
        return;
      }

    } // connectToDB()

    private void dumpData()
    {
      int recs = 0;
      StreamWriter outFile = null;
      string sqlStatement;
      string dateMask;

      // Build the SQL statement
      if (whereClause == "")
      {
        sqlStatement = "select * from " + table;
      }
      else
      {
        sqlStatement = "select * from " + table + " where " + whereClause;
      }
      OracleCommand cmdSQL = new OracleCommand(sqlStatement, oraCon);

      // Read the data
      OracleDataReader dataReader = cmdSQL.ExecuteReader();
      int fieldCount = dataReader.FieldCount;

      // Output SQL*Loader Control file
      try
      {
        outFile = new StreamWriter(table + ".ctl");
        outFile.WriteLine("load data");
        outFile.WriteLine("infile '{0}'", table + ".dat");
        outFile.WriteLine("truncate");
        outFile.WriteLine("into table " + table);
        outFile.WriteLine("fields terminated by '{0}' optionally enclosed by '\"'", delimiter);
        outFile.WriteLine("trailing nullcols");
        outFile.WriteLine("(");
        for (int i = 0; i < fieldCount; i++)
        {
          if (dataReader.GetDataTypeName(i) == DATE)
          {
            dateMask = "        DATE \"MM/DD/YYYY HH12:MI:SS PM\" ";
          }
          else
          {
            dateMask = "";
          }
          if (i == 0)
          {
            outFile.WriteLine("  " + dataReader.GetName(i) + dateMask);
          }
          else
          {
            outFile.WriteLine(", " + dataReader.GetName(i) + dateMask);
          }
        }
        outFile.WriteLine(")");
      }
      catch (Exception ex)
      {
        outputLabel.Text = ex.Message;
      }
      finally
      {
        outFile.Close();
      }

      // Output the data
      try
      {
        outFile = new StreamWriter(table + ".dat");
        while (dataReader.Read())
        {
          recs++;
          for (int i = 0; i < fieldCount; i++)
          {
            try
            {
              outFile.Write(dataReader[i].ToString());
              if ( i < fieldCount - 1)
              {
                outFile.Write(delimiter);
              }
            }
            catch (ArithmeticException)
            {
              outFile.Write( Math.Round((double)dataReader.GetDecimal(i), 2) );
              if ( i < fieldCount - 1)
              {
                outFile.Write(delimiter);
              }
            }
          }
          outFile.WriteLine();
        }
      }
      catch (Exception ex)
      {
        outputLabel.Text = ex.Message;
      }
      finally
      {
        outFile.Close();
      }

      outputLabel.Text = recs + " records written to " + table + ".dat";
      outputLabel.Text += "\nControl file written to " + table + ".ctl";

    } // dumpData()

	} // Class TableDump

}  // namespace TableDump2
