using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;

namespace PlanPro
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private Word.Application wordapp;
        private Word.Documents worddocuments;
        private Word.Document worddocument;
        // private Word.Paragraphs wordparagraphs;
        private Word.Paragraph wordparagraph;
        // кнопка УМР
        private void button1_Click(object sender, EventArgs e)
        {
            YMR y = new YMR(); // переменная для перехода на форму 
            y.Show(); // открытие  формы 
        }

        // кнопка НМНИР
        private void button2_Click(object sender, EventArgs e)
        {
            NMNIR n = new NMNIR(); // переменная для перехода на форму 
            n.Show(); // открытие  формы 
        }
        // кнопка ОМР
        private void button3_Click(object sender, EventArgs e)
        {
            OMR o = new OMR(); // переменная для перехода на форму 
            o.Show(); // открытие  формы 
        }
        // кнопка ВР
        private void button4_Click(object sender, EventArgs e)
        {
            VR v = new VR(); // переменная для перехода на форму 
            v.Show(); // открытие  формы 
        }
        // кнопка сохранить 
        private void button5_Click(object sender, EventArgs e)
        {
            //Закрытие формы
            Application.Exit();
        }
        //кнопка открытия WORD 
        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }


        //Импорт отчёта в MS Word
        private void экспертВMSWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount != 0)
            {
                Double kk1 = GetKK(date_start, date_end, 1);
                //MessageBox.Show(kk1.ToString());
                Double kk2 = GetKK(date_start, date_end, 2);
                //MessageBox.Show(kk2.ToString());
                Double kk = (kk1 + kk2) / 2;
                //Импорт отчёта
                using (DocX document = DocX.Load(@"reports\templates\itogs_kkmp.docx"))
                {
                    try
                    {
                        document.Bookmarks["name_org"].SetText(SettingOrg.getParamOrg(1));
                        document.Bookmarks["period"].SetText(period);
                        document.Bookmarks["kk"].SetText(kk.ToString());
                        document.Bookmarks["kk1"].SetText(kk1.ToString());
                        document.Bookmarks["kk2"].SetText(kk2.ToString());

                    }
                    catch
                    {

                    }



                }
            }
        }
    }

}