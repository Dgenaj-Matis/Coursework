using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using System.IO;
using System.Diagnostics;

namespace Kursovaya_rabota
{
    public partial class Form1 : Form

    {
        DataSet ds;
        MySqlDataAdapter dataAdapter;
        string ConnectString = "Database = zoo_animal2; Data Source = localhost; User Id = id; Password = pass";
        string CommandText = "select * from zoo_animal2.animals";


        public Form1()
        {
            InitializeComponent();

            this.StartPosition = FormStartPosition.CenterScreen;
            this.dataGridView1.BackgroundColor = Color.CornflowerBlue;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;  //устанавливаем полное выделение строки 
            dataGridView1.AllowUserToAddRows = false;   //запрет на ручное добавление строк
            
            button5.Enabled = false;


            using (MySqlConnection connection = new MySqlConnection(ConnectString))
            {
                connection.Open();
                dataAdapter = new MySqlDataAdapter(CommandText, connection);

                ds = new DataSet();

                MySqlCommandBuilder bulder = new MySqlCommandBuilder(dataAdapter);
                dataAdapter.UpdateCommand = bulder.GetUpdateCommand();
                dataAdapter.InsertCommand = bulder.GetInsertCommand();
                dataAdapter.DeleteCommand = bulder.GetDeleteCommand();
                dataAdapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }

        }

        private void Button1_Click(object sender, EventArgs e) //Добавление строки из полей textBox 
        {
            using (MySqlConnection connection = new MySqlConnection(ConnectString))
            {
                if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "")
                {
                    MessageBox.Show("Имеются незаполненные поля!");
                }
                else
                {
                    try
                    {
                        MySqlCommand command = new MySqlCommand("INSERT INTO animals (Помещение,Вид,Кличка,Возраст,Рацион,ФИО)" + "VALUES (@Помещение,@Вид,@Кличка,@Возраст,@Рацион,@ФИО)", connection);
                        command.Parameters.AddWithValue("@Помещение", textBox1.Text);
                        command.Parameters.AddWithValue("@Вид", textBox2.Text);
                        command.Parameters.AddWithValue("@Кличка", textBox3.Text);
                        command.Parameters.AddWithValue("@Возраст", textBox4.Text);
                        command.Parameters.AddWithValue("@Рацион", textBox5.Text);
                        command.Parameters.AddWithValue("@ФИО", textBox6.Text);

                        command.Connection.Open();
                        command.ExecuteNonQuery();
                        textBox1.Clear();
                        textBox2.Clear();
                        textBox3.Clear();
                        textBox4.Clear();
                        textBox5.Clear();
                        textBox6.Clear();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Возможны следующие ошибки:" +
                        "Проверьте правильность типов данных!" +
                        "Проверьте уникальность ключевого поля!");
                    }
                }
            }
            using (MySqlConnection connection = new MySqlConnection(ConnectString)) //обновление dataGridView 
            {
                dataGridView1.DataSource = null;
                connection.Open();
                dataAdapter = new MySqlDataAdapter(CommandText, connection);
                ds = new DataSet();
                MySqlCommandBuilder bulder = new MySqlCommandBuilder(dataAdapter);
                dataAdapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
            }



        }

        private void Button2_Click(object sender, EventArgs e) //Удаление выделенной строки 
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
            }

        }

        private void Button3_Click(object sender, EventArgs e) //Сохранение данных в БД 
        {

                using (MySqlConnection connection = new MySqlConnection(ConnectString))
                {
                dataAdapter.Update(ds.Tables[0]);
                }
            
        }

        int numberRows = 0;

        private void Button4_Click(object sender, EventArgs e) //Запись в Calc
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile ef = new ExcelFile();
            ExcelWorksheet ws = ef.Worksheets.Add("Writing");
            MySqlConnection connection = new MySqlConnection(ConnectString);
            try
            {
                connection.Open();
                MySqlCommand command = new MySqlCommand(CommandText, connection);
                MySqlDataReader reader = command.ExecuteReader();
                int i = 0, j;
                while (reader.Read())
                {
                    for (j = 0; j < 6; j++)
                    {
                        ws.Cells[i, j].Value = reader[j];
                    }
                    i++;
                }
                ws.Cells.GetSubrangeAbsolute(0, 0, i - 1, 6).Sort(false).By(0, false).Apply();
                ef.Save("Zoo.ods");
                connection.Close();
                MessageBox.Show("Данные успешно записаны в файл. Файл находится в проекте.");
                button5.Enabled = true;
                Process.Start("Zoo.ods");

                numberRows = i;
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void Button5_Click(object sender, EventArgs e) //Запись в .txt файл
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                ExcelFile ef = ExcelFile.Load("Zoo.ods");
                ExcelWorksheet ws = ef.Worksheets[0];

                string getStringOutput, tmp;
                int j;
                File.Delete("Zoo.txt");

                StreamWriter writer1 = new StreamWriter("Zoo.txt", true);
                writer1.WriteLine("+---------------------------------------------------------------------------------------------------+");
                writer1.WriteLine("|Номер помещения   " + " Вид животного     " + " Кличка        " + " Возраст  " + " Рацион " + "  Фамилия работника зоопарка" + " |\n");
                writer1.WriteLine("+---------------------------------------------------------------------------------------------------+");
                writer1.Close();

                for (int i = 0; i < numberRows; i++)
                {
                    getStringOutput = "";
                    for (int k = 0; k < 6; k++)
                    {
                        getStringOutput += ws.Cells[i, k].Value.ToString();
                        tmp = ws.Cells[i, k].Value.ToString();

                        switch (k)
                        {
                            case 0:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += " ";
                                    }
                                    break;
                                }

                            case 1:
                                {
                                    for (j = 0; j < 12 - tmp.Length; j++)
                                    {
                                        getStringOutput += " ";
                                    }
                                    break;
                                }
                            case 2:
                                {
                                    for (j = 0; j < 10 - tmp.Length; j++)
                                    {
                                        getStringOutput += " ";
                                    }
                                    break;
                                }
                            case 3:
                                {
                                    for (j = 0; j < 5 - tmp.Length; j++)
                                    {
                                        getStringOutput += " ";
                                    }
                                    break;
                                }
                            case 4:
                                {
                                    for (j = 0; j < 15 - tmp.Length; j++)
                                    {
                                        getStringOutput += " ";
                                    }
                                    break;
                                }
                            case 5:
                                {
                                    for (j = 0; j < 12 - tmp.Length; j++)
                                    {
                                        getStringOutput += " ";
                                    }
                                    break;
                                }
                        }
                        getStringOutput += "\t";
                    }

                    StreamWriter writer = new StreamWriter("Zoo.txt", true);
                    writer.WriteLine("|" + getStringOutput + "    |\n");
                    writer.Close();
                }

                StreamWriter writer2 = new StreamWriter("Zoo.txt", true);
                writer2.WriteLine("+---------------------------------------------------------------------------------------------------+");
                writer2.Close();
                MessageBox.Show("Данные успешно записаны в файл для печати. Файл находится в проекте. ");

                Process.Start("Zoo.txt");
                
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

       
    }
}
