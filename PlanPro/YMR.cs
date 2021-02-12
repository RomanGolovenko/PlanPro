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
using Xceed.Words.NET;

namespace PlanPro
{
    public partial class YMR : Form
    {

        // Глобальные переменные
        // Точка в хода в бд
        string connStr = "server = osp74.ru; user = st_2_5;database = st_2_5;password = 29259054; port =33333";// Строка входа в бд
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
        //Запрос для вывода строк в БД
        string commandStr = "SELECT TypeWork AS 'Вид работы', ReportForm AS 'Форма отчетности', Deadline AS 'Срок выполнения ', Hours AS 'Обьем часов', Mark AS 'Отметка о выполнении' FROM tabYMR";


        public YMR()
        {
            InitializeComponent();
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
            dataGridView1.Columns[2].Visible = true;
            dataGridView1.Columns[3].Visible = true;
            dataGridView1.Columns[4].Visible = true;

            //Ширина полей
            dataGridView1.Columns[0].FillWeight = 90;
            dataGridView1.Columns[1].FillWeight = 50;
            dataGridView1.Columns[2].FillWeight = 50;
            dataGridView1.Columns[3].FillWeight = 20;
            dataGridView1.Columns[4].FillWeight = 20;


            //Растягивание полей грида
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            //Убираем заголовки строк
            dataGridView1.RowHeadersVisible = false;
            //Показываем заголовки столбцов
            dataGridView1.ColumnHeadersVisible = true;
        }
        //Метод наполнения DataGreed
        public void GetListUsers()
        {
            
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
            dataGridView1.ReadOnly = false;
            button1.Enabled =true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

           
           
          
           
        }
        

            //кнопка сохранить
            private void button1_Click(object sender, EventArgs e)
        {
            MyDA.UpdateCommand = new MySqlCommand(commandStr, conn_db);
            MyDA.Update(table);
            dataGridView1.ReadOnly = true;
            button1.Enabled = true;
            
            ////Закрытие формы
            //Close();
        }
        // метод обработки исключений грида, когда пользователь вводит не целочисленые
        // значения в столбец "количество часов"
        private void dataGridView1_DataError_1(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show(" Неверный формат введенных данных  ");
        }

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    MenuItem_Click();
        //}

        

        //private void MenuItem_Click()
        //{
        //    string a = "переменая";
        //    if (dataGridView1.RowCount != 0)
        //    {
        //        using (var document = DocX.Load(@"C:\Users\user\source\repos\PlanPro-master\kkmp.docx"))
        //        {
        //            //document.Bookmarks["a"].SetText(a);

        //            var criterias = document.AddTable(dataGridView1.RowCount + 2, dataGridView1.ColumnCount - 1);
        //            //Озаглавливание полей                    
        //            criterias.Rows[1].Cells[0].Paragraphs[0].Append(dataGridView1.Columns[1].HeaderText);
        //            criterias.Rows[1].Cells[1].Paragraphs[0].Append("Всего");
        //            criterias.Rows[1].Cells[2].Paragraphs[0].Append("I");
        //            criterias.Rows[1].Cells[3].Paragraphs[0].Append("II");
        //            //Объединение ячеек в 0 строке
        //            criterias.Rows[0].MergeCells(1, 3);
        //           //criterias.Rows[0].MergeCells(2, 4);
        //            //Наименование объединенных ячеек
        //            criterias.Rows[0].Cells[1].Paragraphs[0].Append("Количество экспертиз");
        //            //criterias.Rows[0].Cells[2].Paragraphs[0].Append("Количество дефектов");
        //            //Цикл заполнения таблицы с критериями
        //            for (int i = 2; i < dataGridView1.RowCount + 2; i++)
        //            {
        //                for (int j = 1; j < dataGridView1.ColumnCount; j++)
        //                {
        //                    string value = dataGridView1[j, i - 2].Value.ToString();
        //                    criterias.Rows[i].Cells[j - 1].Paragraphs[0].Append(value);

        //                }
                        
        //            }
                    
        //            //Настройка внешнего вида таблицы                    
        //            criterias.Design = TableDesign.TableGrid;
        //            criterias.Alignment = Alignment.center;
        //            //Вставка таблицы в документ
        //            document.InsertTable(criterias);

        //        }
        //    }
        //}

      
    }
    
}
