using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace DiplomDeskTop
{
    public partial class Form1 : Form
    {
        StartForm StartForm;
        public string table;
        DataSet ds;
        SqlDataAdapter adapter;
        SqlCommandBuilder commandBuilder;
        public const string sql = "SELECT * FROM ";
        int chosencolomn;
        int ChosenRow;
        string lastString;
        string curentString;
        public Form1(StartForm StartForm)
        {
            try
            {
                this.StartForm = StartForm;
                InitializeComponent();
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView1.AllowUserToAddRows = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                StartForm.Enabled = true;
                StartForm.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (StartForm.connection.State == ConnectionState.Open)
                {
                    TableChouse tableChouse = new TableChouse(StartForm.connection, this);
                    tableChouse.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }
        public void DataShow(string str)
        {
            try
            {                 
                if (curentString != null)
                {
                    lastString = curentString;
                }
                curentString = str;
                adapter = new SqlDataAdapter(str, StartForm.connection);
                ds = new DataSet();
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }  
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                commandBuilder = new SqlCommandBuilder(adapter);
                adapter.Update(ds);
                MessageBox.Show("Готово!");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                // проверяем выбрана ли таблица
                if (dataGridView1.Rows.Count < 1)
                {
                    MessageBox.Show("Таблица не выбрана");
                    return;
                }
                //Выбираем файл
                openFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx";
                openFileDialog1.FileName = "";
                if (openFileDialog1.ShowDialog() == DialogResult.Cancel) return;
                string filename = openFileDialog1.FileName;
                // для eppplus нужно выставить лицензию
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                //создадим виртуальный excel
                ExcelPackage excel = new ExcelPackage();
                var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
                //заполним название колон и сделаем им авторасширение (оно не сработало)
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    workSheet.Column(i + 1).AutoFit();
                    if (dataGridView1.Columns[i].Name == null)
                        continue;
                    workSheet.Cells[1, i + 1].Value = (dataGridView1.Columns[i].Name).ToString();
                }
                //заполним таблицу
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[i, j].Value == null)
                            continue;
                        workSheet.Cells[j + 2, i + 1].Value = (dataGridView1[i, j].Value).ToString();
                    }
                }
                //выгрузим в выбранный excel
                File.WriteAllBytes(filename, excel.GetAsByteArray());
                MessageBox.Show("Успешно!");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //сортировка правой кнопкой мыши и контекстным меню
            try 
            {
                if (e.Button == MouseButtons.Right)
                {
                    contextMenuStrip1.Show(Control.MousePosition);
                    chosencolomn = e.ColumnIndex;
                    ChosenRow = e.RowIndex;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
}

        private void SortByColomn_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewColumn colomn = dataGridView1.Columns[chosencolomn];
                if (dataGridView1.SortedColumn == colomn & dataGridView1.SortOrder == System.Windows.Forms.SortOrder.Ascending)
                {
                    dataGridView1.Sort(colomn, ListSortDirection.Descending);
                    return;
                }
                dataGridView1.Sort(colomn, ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SurchByCell_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1[chosencolomn, ChosenRow].Value == null)
                    return;
                string colomnname = dataGridView1.Columns[chosencolomn].Name;
                if (table.StartsWith("select [Номер_договора] as 'Номер договора',[Наименование] as 'Тип отходов',SUM([Изменение_накопления]) as 'Накопение (кг)'"))
                {
                    string table2 = "select [Номер_договора] as 'Номер договора',[Наименование] as 'Тип отходов',SUM([Изменение_накопления]) as 'Накопение (кг)' from[dbo].[Клиенты] K inner join[dbo].[Накопление_клиентов] N on K.[Id_клиента] = N.[Id_клиента] inner join[dbo].[Тип_отходов] T on N.Id_Типа = T.Id_Типа";
                    switch (colomnname)
                    {
                        case "Номер договора":
                            table2 += (" WHERE K.[Номер_договора]" + " = N\'" + dataGridView1[chosencolomn, ChosenRow].Value.ToString() + "\' group by [Номер_договора],[Наименование]  order by [Номер_договора]");
                            DataShow(table2);
                            break;
                        case "Тип отходов":
                            table2 += (" WHERE t.[Наименование]" + " = N\'" + dataGridView1[chosencolomn, ChosenRow].Value.ToString() + "\' group by [Номер_договора],[Наименование]  order by [Номер_договора]");
                            DataShow(table2);
                            break;
                        case "Накопение (кг)":
                            table2 = ("select * from (select [Номер_договора] as 'Номер договора',[Наименование] as 'Тип отходов', SUM([Изменение_накопления]) as 'Накопение (кг)' from [dbo].[Клиенты] K inner join [dbo].[Накопление_клиентов] N on K.[Id_клиента] = N.[Id_клиента] inner join [dbo].[Тип_отходов] T on N.Id_Типа = T.Id_Типа group by [Номер_договора], [Наименование]) as tab where tab.[Накопение (кг)] = " + dataGridView1[chosencolomn, ChosenRow].Value.ToString().Replace(',','.') + " order by tab.[Номер договора]");
                            DataShow(table2);
                            break;
                        default:
                            MessageBox.Show(colomnname);
                            break;
                    }
                }
                else
                {
                    DataShow(table + " WHERE " + colomnname + " = N\'" + dataGridView1[chosencolomn, ChosenRow].Value.ToString() + "\'");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            table = "select [Номер_договора] as 'Номер договора',[Наименование] as 'Тип отходов',SUM([Изменение_накопления]) as 'Накопение (кг)' from [dbo].[Клиенты] K inner join [dbo].[Накопление_клиентов] N on K.[Id_клиента] = N.[Id_клиента] inner join [dbo].[Тип_отходов] T on N.Id_Типа = T.Id_Типа group by [Номер_договора],[Наименование]  order by [Номер_договора]";
            DataShow(table);
        }

        private void Refresh_Click(object sender, EventArgs e)
        {
            if (lastString != null)
                DataShow(curentString);
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            if (lastString != null)
                DataShow(lastString);
        }
    }
}
