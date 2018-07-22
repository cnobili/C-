/*
 * GetData.cs
 *
 * Program that writes out the result set of a SQL query
 * to a delimited flat file.
 *
 * Craig Nobili
 */
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.Odbc;
using System.IO;

namespace GetData
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class GetData: System.Windows.Forms.Form
	{
    private System.Windows.Forms.Label odbcLabel;
    private System.Windows.Forms.Label userLabel;
    private System.Windows.Forms.Label passwdLabel;
    private System.Windows.Forms.Label sqlLabel;
    private System.Windows.Forms.Label outputFileLabel;
    private System.Windows.Forms.Label DelimiterLabel;
    private System.Windows.Forms.TextBox odbcTextBox;
    private System.Windows.Forms.TextBox userTextBox;
    private System.Windows.Forms.TextBox passwdTextBox;
    private System.Windows.Forms.TextBox sqlTextBox;
    private System.Windows.Forms.TextBox outputFileTextBox;
    private System.Windows.Forms.TextBox delimiterTextBox;
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.Container components = null;

    private const string DATE     = "Date";
    private string odbcDsn;
    private string userid;
    private string passwd;
    private string sql;
    private string outputFile;
    private System.Windows.Forms.Button submitButton;
    private System.Windows.Forms.Label outputLabel;
    private string delimiter;
    private OdbcConnection con;

		public GetData()
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
      this.odbcLabel = new System.Windows.Forms.Label();
      this.userLabel = new System.Windows.Forms.Label();
      this.passwdLabel = new System.Windows.Forms.Label();
      this.sqlLabel = new System.Windows.Forms.Label();
      this.outputFileLabel = new System.Windows.Forms.Label();
      this.DelimiterLabel = new System.Windows.Forms.Label();
      this.odbcTextBox = new System.Windows.Forms.TextBox();
      this.userTextBox = new System.Windows.Forms.TextBox();
      this.passwdTextBox = new System.Windows.Forms.TextBox();
      this.sqlTextBox = new System.Windows.Forms.TextBox();

      // Set the Multiline property to true.
      this.sqlTextBox.Multiline = true;
      // Add vertical scroll bars to the TextBox control.
      this.sqlTextBox.ScrollBars = ScrollBars.Vertical;
      // Allow the RETURN key to be entered in the TextBox control.
      this.sqlTextBox.AcceptsReturn = true;
      // Allow the TAB key to be entered in the TextBox control.
	    this.sqlTextBox.AcceptsTab = true;
	    // Set WordWrap to true to allow text to wrap to the next line.
	    this.sqlTextBox.WordWrap = true;

      this.outputFileTextBox = new System.Windows.Forms.TextBox();
      this.delimiterTextBox = new System.Windows.Forms.TextBox();
      this.submitButton = new System.Windows.Forms.Button();
      this.outputLabel = new System.Windows.Forms.Label();
      this.SuspendLayout();
      //
      // odbcLabel
      //
      this.odbcLabel.Location = new System.Drawing.Point(16, 40);
      this.odbcLabel.Name = "odbcLabel";
      this.odbcLabel.Size = new System.Drawing.Size(112, 23);
      this.odbcLabel.TabIndex = 0;
      this.odbcLabel.Text = "ODBC DSN";
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
	    // DelimiterLabel
	    //
	    this.DelimiterLabel.Location = new System.Drawing.Point(16, 136);
	    this.DelimiterLabel.Name = "DelimiterLabel";
	    this.DelimiterLabel.TabIndex = 3;
      this.DelimiterLabel.Text = "Delimiter";
      //
      // outputFileLabel
      //
      this.outputFileLabel.Location = new System.Drawing.Point(16, 168);
      this.outputFileLabel.Name = "outputFileLabel";
      this.outputFileLabel.TabIndex = 4;
      this.outputFileLabel.Text = "Output File";
      //
      // sqlLabel
      //
      this.sqlLabel.Location = new System.Drawing.Point(16, 232);
      this.sqlLabel.Name = "sqlLabel";
      this.sqlLabel.TabIndex = 6;
      this.sqlLabel.Text = "SQL";
      //
      // odbcTextBox
      //
      this.odbcTextBox.Location = new System.Drawing.Point(120, 40);
      this.odbcTextBox.Name = "odbcTextBox";
      this.odbcTextBox.TabIndex = 7;
      this.odbcTextBox.Text = "";
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
      // delimiterTextBox
      //
      this.delimiterTextBox.Location = new System.Drawing.Point(120, 136);
      this.delimiterTextBox.Name = "delimiterTextBox";
      this.delimiterTextBox.TabIndex = 10;
      this.delimiterTextBox.Text = "";
      //
      // outputFileTextBox
      //
      this.outputFileTextBox.Location = new System.Drawing.Point(120, 168);
      this.outputFileTextBox.Name = "outputFileTextBox";
      this.outputFileTextBox.TabIndex = 11;
      this.outputFileTextBox.Text = "";
      //
      // sqlTextBox
      //
      this.sqlTextBox.Location = new System.Drawing.Point(120, 232);
      this.sqlTextBox.Name = "sqlTextBox";
      this.sqlTextBox.TabIndex = 13;
      this.sqlTextBox.Text = "";
      this.sqlTextBox.Width = 650;
      this.sqlTextBox.Height = 400;
      //
      // submitButton
      //
      this.submitButton.Location = new System.Drawing.Point(24, 650);
      this.submitButton.Name = "submitButton";
      this.submitButton.TabIndex = 14;
      this.submitButton.Text = "Submit";
      this.submitButton.Click += new System.EventHandler(this.submitButton_Click);
      //
      // outputLabel
      //
      this.outputLabel.Location = new System.Drawing.Point(628, 725);
      this.outputLabel.Name = "outputLabel";
      this.outputLabel.Size = new System.Drawing.Size(224, 150);
      this.outputLabel.TabIndex = 15;
      //
      // GetData
      //
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.ClientSize = new System.Drawing.Size(824, 813);
      this.Controls.Add(this.outputLabel);
      this.Controls.Add(this.submitButton);
      this.Controls.Add(this.delimiterTextBox);
      this.Controls.Add(this.outputFileTextBox);
      this.Controls.Add(this.sqlTextBox);
      this.Controls.Add(this.passwdTextBox);
      this.Controls.Add(this.userTextBox);
      this.Controls.Add(this.odbcTextBox);
      this.Controls.Add(this.DelimiterLabel);
      this.Controls.Add(this.outputFileLabel);
      this.Controls.Add(this.sqlLabel);
      this.Controls.Add(this.passwdLabel);
      this.Controls.Add(this.userLabel);
      this.Controls.Add(this.odbcLabel);
      this.Name = "GetData";
      this.Text = "GetData";
      this.Load += new System.EventHandler(this.GetData_Load);
      this.ResumeLayout(false);

    }
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.Run(new GetData());
		}

    private void GetData_Load(object sender, System.EventArgs e)
    {
    }

    private void submitButton_Click(object sender, System.EventArgs e)
    {
      odbcDsn = odbcTextBox.Text;
      userid = userTextBox.Text;
      passwd = passwdTextBox.Text;
      sql = sqlTextBox.Text;
      outputFile = outputFileTextBox.Text;
      delimiter = delimiterTextBox.Text;

      // Connect to database and get the data
      connectToDB();
      dumpData();
      con.Close();
      con.Dispose();
    }

    private void connectToDB()
    {
      string connStr = "DSN=" + odbcDsn + ";uid=" + userid + ";pwd=" + passwd + ";";

      // Connect to database
      con = new OdbcConnection(connStr);
      try
      {
        con.Open();
      }
      catch (Exception ex)
      {
        outputLabel.Text = "Failed to Connect to database:" + ex.Message;
        return;
      }

    } // connectToDB()

    private void dumpData()
    {
      String rec = null;
      String fieldSep = null;
      int recs = 0;
      StreamWriter outFile = null;
      string sqlStatement;

      // Build the SQL statement
      sqlStatement = sql;

      OdbcCommand cmdSQL = new OdbcCommand(sqlStatement, con);

      // Read the data
      OdbcDataReader dataReader = cmdSQL.ExecuteReader();
      int fieldCount = dataReader.FieldCount;

      // Output the data
      try
      {
        outFile = new StreamWriter(outputFile);
                
        rec = "";
        fieldSep = "";
        for (int i = 0; i < fieldCount; i++)
        {
          rec += fieldSep;
          rec += dataReader.GetName(i);
          fieldSep = delimiter;
        }
        outFile.WriteLine(rec);
                
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

      outputLabel.Text = recs + " records written to " + outputFile;

    } // dumpData()

	} // Class GetData

}  // namespace GetData
