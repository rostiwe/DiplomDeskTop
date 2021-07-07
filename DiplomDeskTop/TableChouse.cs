using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace DiplomDeskTop
{
    public partial class TableChouse : Form
    {
        SqlConnection SqlConnection;
        Form1 Form1;
        


        public TableChouse(SqlConnection SqlConnection, Form1  Form1)
        {
            this.Form1 = Form1;
            this.SqlConnection = SqlConnection;
            InitializeComponent();
            List<string> TableNames = GetTables(SqlConnection);
            for (int i = 0; i< TableNames.Count; i++)
            {
                comboBox1.Items.Add(TableNames[i]);
            }
            Form1.Enabled = false;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                Form1.table = "SELECT * FROM "+comboBox1.Text;
                Form1.DataShow(Form1.table);
                this.Close();
            }

        }
        public static List<string> GetTables(SqlConnection SqlConnection)
        {
            DataTable schema = SqlConnection.GetSchema("Tables");
            List<string> TableNames = new List<string>();
            foreach (DataRow row in schema.Rows)
            {
                TableNames.Add(row[2].ToString());
            }
            return TableNames;
        }

        private void TableChouse_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1.Enabled = true;
        }
    }
}
