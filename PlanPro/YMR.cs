using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using MySql.Data;

namespace PlanPro
{
    public partial class YMR : Form
    {

        // Глобальные переменные
        // Точка в хода в бд
        string connStr = "server = osp74.ru; user = st_2_1;database = st_2_1;password = 33674763; port =33333";// Строка входа в бд
        MySqlConnection conn_db;
        private MySqlDataAdapter MyDA = new MySqlDataAdapter();
        //Объявление BindingSource, основная его задача, это обеспечить унифицированный доступ к источнику данных.
        private BindingSource bSource = new BindingSource();
        //DataSet - расположенное в оперативной памяти представление данных, обеспечивающее согласованную реляционную программную 
        //модель независимо от источника данных.DataSet представляет полный набор данных, включая таблицы, содержащие, упорядочивающие 
        //и ограничивающие данные, а также связи между таблицами.
        private DataSet ds = new DataSet();
        //Представляет одну таблицу данных в памяти.
        private DataTable table = new DataTable();

        public YMR()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Закрытие формы
            Close();
        }

        private void YMR_Load(object sender, EventArgs e)
        {
            //Инициализируем соединение с БД
            conn_db = new MySqlConnection(connStr);
            //Вызываем метод для заполнение дата Грида
            GetListUsers();
            //Видимость полей в гриде
            dataGridView1.Columns[0].Visible = true;
            dataGridView1.Columns[1].Visible = true;
            //dataGridView1.Columns[2].Visible = true;
            //dataGridView1.Columns[3].Visible = true;
            //dataGridView1.Columns[4].Visible = true;

            //Ширина полей
            dataGridView1.Columns[0].FillWeight = 10;
            dataGridView1.Columns[1].FillWeight = 50;
            //dataGridView1.Columns[2].FillWeight = 50;
            //dataGridView1.Columns[3].FillWeight = 50;
            //dataGridView1.Columns[4].FillWeight = 50;


            //Растягивание полей грида
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //Убираем заголовки строк
            dataGridView1.RowHeadersVisible = false;
            //Показываем заголовки столбцов
            dataGridView1.ColumnHeadersVisible = true;
        }
        //Метод наполнения DataGreed
        public void GetListUsers()
        {
            //Запрос для вывода строк в БД
            string commandStr = "SELECT id_categ AS 'ID', title_categ AS 'title_categ' FROM t_categ";
            //Открываем соединение
            conn_db.Open();
            //Объявляем команду, которая выполнить запрос в соединении conn
            MyDA.SelectCommand = new MySqlCommand(commandStr, conn_db);
            //Заполняем таблицу записями из БД
            MyDA.Fill(table);
            //Указываем, что источником данных в bindingsource является заполненная выше таблица
            bSource.DataSource = table;
            //Указываем, что источником данных ДатаГрида является bindingsource 
            dataGridView1.DataSource = bSource;
            //Закрываем соединение
            conn_db.Close();

        }
    }
}
