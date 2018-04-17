using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace IFE
{

    public partial class Form1 : Form
    {
        string cons = "", conStr = "";
        static string selectCommand = "SELECT * FROM [dbo].[standarts]"; 
        static string selectCommand1 = "SELECT * FROM [dbo].[import]";
        string insertCommand = "";
        SqlConnection cn = null;
        SqlDataAdapter dataAdapter = null;
        SqlDataAdapter dataAdapter1 = null;
        SqlCommand com1 = null;
        SqlCommandBuilder bldr = null;
        SqlCommandBuilder bldr1 = null;
        System.Data.DataTable dataTable = null;
        System.Data.DataTable dataTable1 = null;
        int counter = 3, rowCounter = 1;



        public Form1()
        {

            InitializeComponent();
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //Подключение к БД
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                cons = openFileDialog1.FileName;

                conStr = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + cons + "";


                cn = new SqlConnection(); //Подключение SQL
                //cn.ConnectionString = conStr;

                using (cn)
                {

                    cn.ConnectionString = conStr; //Строка подключения к бд
                    try
                    {
                        dataAdapter = new SqlDataAdapter(); //Инициализация адаптера для источника данных
                        bldr = new SqlCommandBuilder(dataAdapter); //Инициализация построителя команд SQL, основываясь на подключенном адаптере

                        SqlCommand sqlCommand = new SqlCommand(selectCommand, cn);//Инициализация команды SELECT с использования подключения SQL
                        dataAdapter.SelectCommand = sqlCommand; //Присвоение адаптеру готовой Select команды с подключением 
                        dataAdapter.InsertCommand = bldr.GetInsertCommand(); //Автогенерация команды insert(создание новой строки БД) и присвоение адаптеру
                        dataAdapter.UpdateCommand = bldr.GetUpdateCommand(); //Автогенерация команды update(обновление существующей строки) и присвоение адаптеру

                        cn.Open(); //Открытие подключения

                        dataTable = new System.Data.DataTable(); //Инициализируем табличное представление
                        dataAdapter.Fill(dataTable); //Заполняем табличное представление значениями из адаптера
                        dataGridView1.DataSource = dataTable; //Присваиваем табличное представление(таблицу) как источник данных для dataGridView1
                        dataGridView1.Columns[0].Visible = false;

                        //Заполняется верхний dataGridView1 из таблицы standarts

                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK);
                    }
                    finally
                    {
                        cn.Close();//Закрытие подключения
                    }

                }

                using (cn)
                {
                    //Все по аналогии с dataGridView1
                    cn.ConnectionString = conStr;
                    try
                    {
                        cn.Open();
                        dataAdapter1 = new SqlDataAdapter();
                        bldr1 = new SqlCommandBuilder(dataAdapter1);

                        SqlCommand sqlCommand = new SqlCommand(selectCommand1, cn);
                        dataAdapter1.SelectCommand = sqlCommand;
                        dataAdapter1.InsertCommand = bldr1.GetInsertCommand();
                        dataAdapter1.UpdateCommand = bldr1.GetUpdateCommand();



                        dataTable1 = new System.Data.DataTable();
                        dataAdapter1.Fill(dataTable1);
                        dataGridView2.DataSource = dataTable1;
                        dataGridView2.Columns[0].Visible = false;

                        //По тому же принципу заполняется нижний dataGridView2 из таблицы import


                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK);
                    }
                    finally
                    {
                        cn.Close();
                    }

                }
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            //Считывание данных из .xls
            counter = 3;
            if (openFileDialog1.ShowDialog() == DialogResult.OK) //Диалоговое окно, которое срабатывает при выборе файла и нажатии ок
            {
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application(); //Открывает экземпляр Excel

                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);//Открытие файла с книгой

                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;// Инициализация  страницы
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[2];// Выбор конкретной лист

                for (int i = 1; i < 3; i++)
                {

                    Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("C" + Convert.ToString(i + 1), "C" + Convert.ToString(i + 1)); //Считывание данных с конкретной ячейки листа

                    dataGridView1.Rows[0].Cells[i].Value = range.Text;// Записываем полученное значение ячейки в dataGridView1

                }

                for (int i = 1; i < 11; i++)
                {

                    if (i == 3 || i == 5 || i == 7 || i == 8 || i == 9)
                    {
                        continue;
                    }
                    else
                    {
                        Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("F" + Convert.ToString(i + 10), "F" + Convert.ToString(i + 10));
                        dataGridView1.Rows[0].Cells[counter].Value = range.Text;
                        counter++;
                    }

                }

                dataGridView2.Rows[0].Cells[0].Value = 1;

                int sheetsCount = ObjWorkBook.Sheets.Count;

                for (int j = 3; j < (Convert.ToInt32(textBox1.Text)+3); j++)
                {
                    counter = 4;

                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[j]; //Выбор листов в цикле


                    for (int i = 1; i < 4; i++)
                    {

                        Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("C" + Convert.ToString(i + 1), "C" + Convert.ToString(i + 1));

                        dataGridView2.Rows[rowCounter - 1].Cells[i].Value = range.Text;

                    }

                    for (int i = 1; i < 11; i++)
                    {

                        if (i == 2 || i == 4 || i == 6 || i == 8 || i == 9)
                        {
                            continue;
                        }
                        else
                        {
                            if (i == 5)
                            {
                                Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("H" + Convert.ToString(i + 10), "H" + Convert.ToString(i + 10));
                                dataGridView2.Rows[rowCounter - 1].Cells[counter].Value = range.Text;
                                counter++;
                            }
                            else
                            {
                                Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("G" + Convert.ToString(i + 10), "G" + Convert.ToString(i + 10));
                                if (range.Text != "-")
                                {
                                    dataGridView2.Rows[rowCounter - 1].Cells[counter].Value = range.Text;
                                    counter++;
                                }
                                else
                                {
                                    dataGridView2.Rows[rowCounter - 1].Cells[counter].Value = 0;
                                }
                            }
                        }

                        

                    }
                    SqlCommand com1 = new SqlCommand(); //Инициализация экземпляра команды SQL
                    using (SqlConnection cn = new SqlConnection())// Инициализация экземпляра подключения
                    {
                        cn.ConnectionString = conStr;//Приствоение экземпяляру подключения соответствующей строки подключения
                        cn.Open();//Открытие подключения
                        com1.CommandText = String.Format("INSERT INTO [dbo].[import]([Название], [Партия], [Дата испытаний], [Вес, результат], [Емкость, Ф, результат], [ESR, мОм, результат], [КПД, %, результат], [Температура, С, результат] ) " +
                            "VALUES ( '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}')", dataGridView2.Rows[rowCounter].Cells[1].Value, dataGridView2.Rows[rowCounter].Cells[2].Value, dataGridView2.Rows[rowCounter].Cells[3].Value, 
                            dataGridView2.Rows[rowCounter].Cells[4].Value, dataGridView2.Rows[rowCounter].Cells[5].Value, 
                            dataGridView2.Rows[rowCounter].Cells[6].Value, dataGridView2.Rows[rowCounter].Cells[7].Value, 
                            dataGridView2.Rows[rowCounter].Cells[8].Value);
                        //Использование не сохраненных данных, записанных в dataGridview2 на предыдущем шаге для создания строки ввода(INSERT)
                        //
                        com1.Connection = cn;
                        com1.ExecuteNonQuery();

                        cn.Close();
                    }

                    rowCounter++;
                    //counter = 3;



                }
                ObjExcel.Quit();// Закрыть открытый экземпляр Excel
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {

            using (SqlConnection cn = new SqlConnection(conStr)) //инициализация подключения SQL с присвоением строки подключения к БД
            {
                SqlCommand sqlCom = new SqlCommand(selectCommand, cn); //Инициализация команды SELECT sql для таблицы standarts
                SqlCommand sqlCom1 = new SqlCommand(selectCommand1, cn); //Инициализация команды SELECT sql для таблицы import

                dataAdapter = new SqlDataAdapter(selectCommand, cn); //Переопределяется адаптер для таблицы standarts
                dataAdapter1 = new SqlDataAdapter(selectCommand1, cn); //Переопределяется адаптер для таблицы import

                SqlCommandBuilder builder = new SqlCommandBuilder(dataAdapter); //Инициализация построителя команд для адаптера таблицы standarts
                SqlCommandBuilder builder1 = new SqlCommandBuilder(dataAdapter1); //Инициализация построителя команд для адаптера таблицы import

                dataAdapter.SelectCommand = sqlCom; //Присвоение ранее созданной команды Select адаптеру
                dataAdapter.InsertCommand = builder.GetInsertCommand(); //Присвоение сгенерированной команды INSERT адаптеру таблицы standarts
                dataAdapter.UpdateCommand = builder.GetUpdateCommand(); // Присвоение сгенерированной команды UPDATE адаптеру таблицы standarts

                dataAdapter1.SelectCommand = sqlCom1; //Присвоение ранее созданной команды Select адаптеру
                dataAdapter.InsertCommand = builder1.GetInsertCommand(); //Присвоение сгенерированной команды INSERT адаптеру таблицы import
                dataAdapter.UpdateCommand = builder.GetUpdateCommand(); // Присвоение сгенерированной команды UPDATE адаптеру таблицы import

                dataAdapter.Update(dataTable); //Вызывает соответствующие инструкции INSERT, UPDATE или DELETE для каждой вставки, обновления или удаления строки в указанном DataTable standarts
                dataAdapter1.Update(dataTable1); //Вызывает соответствующие инструкции INSERT, UPDATE или DELETE для каждой вставки, обновления или удаления строки в указанном DataTable import
                cn.Close(); //закрытие подключения

            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            for (int i = 0; dataGridView1.Rows[0].Cells[1].Value.ToString() == dataGridView2.Rows[i].Cells[1].Value.ToString(); i++)
            {
               
                if (Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value) < Convert.ToInt32(dataGridView1.Rows[0].Cells[3].Value))
                    dataGridView2.Rows[i].Cells[4].Style.BackColor = System.Drawing.Color.Green;
                else
                    dataGridView2.Rows[i].Cells[4].Style.BackColor = System.Drawing.Color.Red;

                if (Convert.ToInt32(dataGridView2.Rows[i].Cells[5].Value) == Convert.ToInt32(dataGridView1.Rows[0].Cells[4].Value))
                    dataGridView2.Rows[i].Cells[5].Style.BackColor = System.Drawing.Color.Green;
                else
                    dataGridView2.Rows[i].Cells[5].Style.BackColor = System.Drawing.Color.Red;

                if (Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value) < Convert.ToInt32(dataGridView1.Rows[0].Cells[5].Value))
                    dataGridView2.Rows[i].Cells[6].Style.BackColor = System.Drawing.Color.Green;
                else
                    dataGridView2.Rows[i].Cells[6].Style.BackColor = System.Drawing.Color.Red;

                if (Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value) > Convert.ToInt32(dataGridView1.Rows[0].Cells[6].Value))
                    dataGridView2.Rows[i].Cells[7].Style.BackColor = System.Drawing.Color.Green;
                else
                    dataGridView2.Rows[i].Cells[7].Style.BackColor = System.Drawing.Color.Red;

                if (Convert.ToInt32(dataGridView2.Rows[i].Cells[8].Value) < Convert.ToInt32(dataGridView1.Rows[0].Cells[7].Value))
                    dataGridView2.Rows[i].Cells[8].Style.BackColor = System.Drawing.Color.Green;
                else
                    dataGridView2.Rows[i].Cells[8].Style.BackColor = System.Drawing.Color.Red;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

    }
}
