using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlanPro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string FIO_prep = textBox1.Text; // ФИО препода с формы регистрации
            string PCK = comboBox1.Text; // ФИО ПЦК с формы регистрации


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && comboBox1.Text == "")
            {
                MessageBox.Show(" Заполните поля ");

            }

        }
    }
}
