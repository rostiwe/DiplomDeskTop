using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using System.Data.SqlClient;
using Microsoft.Data.SqlClient;

namespace DiplomDeskTop
{
    public partial class StartForm : Form
    {
        //string server;
        public StartForm()
        {
            InitializeComponent();
            comboBox1.Text = comboBox1.Items[0].ToString();
            //server = "diplommusortry1.database.windows.net";
        }
        public string connectionString;
        public SqlConnection connection = new SqlConnection();
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != comboBox1.Items[0].ToString()) {
                bool mark = false;
                if (textBox1.Text == "")
                {
                    mark = true;
                    label4.Visible = true;
                }
                else label4.Visible = false;
                if (textBox2.Text == "")
                {
                    mark = true;
                    label5.Visible = true;
                }
                else label5.Visible = false;
                if (mark)
                {
                    return;
                }
                Connect("Data Source=diplommusortry1.database.windows.net;Initial Catalog=diplom;User ID=" + textBox1.Text + ";Password=" + textBox2.Text + ";Connect Timeout=60;Encrypt=True;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
                if (connection.State == ConnectionState.Open)
                {
                    openform();
                }
            }
            else
            {
                //Connect("Server="+server+";Database=Diplom;Trusted_Connection=True;");
                Connect("Data Source=diplommusortry1.database.windows.net;Initial Catalog=diplom;Trusted_Connection=True;Connect Timeout=60;Encrypt=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
                if (connection.State == ConnectionState.Open)
                {
                    openform();
                }    
            }
        }
        void openform ()
        {
            Form1 Form1 = new Form1(this);
            Form1.Show();
            this.Enabled = false;
            Visible = false;
        }
        void Connect (string text)
        {
            if (connection.State == ConnectionState.Closed)
            {
                connection.ConnectionString = text;
                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            
            if (comboBox1.Text == comboBox1.Items[0].ToString())
            {
                textBox1.Enabled = false;
                textBox2.Enabled = false;
            }
            else
            {
                textBox1.Enabled = true;
                textBox2.Enabled = true;
            }
        }
    }
}
