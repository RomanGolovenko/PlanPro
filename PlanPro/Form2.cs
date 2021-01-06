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
            Close();
        }
        //кнопка открытия WORD 
        private void button6_Click(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(((Button)(sender)).Tag);
            switch (i)
            {
                case 1:

                    worddocument.Content.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    worddocument.Content.ParagraphFormat.LeftIndent =
                     worddocument.Content.Application.CentimetersToPoints((float)2);
                    worddocument.Content.ParagraphFormat.RightIndent =
                     worddocument.Content.Application.CentimetersToPoints((float)1);

                    //Вставляем в документ 4 параграфа
                    object oMissing = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissing);
                    worddocument.Paragraphs.Add(ref oMissing);
                    worddocument.Paragraphs.Add(ref oMissing);
                    worddocument.Paragraphs.Add(ref oMissing);
                    //Переходим к первому добавленному параграфу
                    wordparagraph = worddocument.Paragraphs[2];
                    Word.Range wordrange = wordparagraph.Range;
                    //Добавляем таблицу в начало второго параграфа
                    Object defaultTableBehavior =
                     Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior =
                     Word.WdAutoFitBehavior.wdAutoFitWindow;
                    Word.Table wordtable1 = worddocument.Tables.Add(wordrange, 5, 5,
                      ref defaultTableBehavior, ref autoFitBehavior);
                    //Сдвигаемся вниз в конец документа
                    object unit;
                    object extend;
                    unit = Word.WdUnits.wdStory;
                    extend = Word.WdMovementType.wdMove;
                    wordapp.Selection.EndKey(ref unit, ref extend);
                    //Вставляем таблицу по месту курсора
                    Word.Table wordtable2 = worddocument.Tables.Add(
                      wordapp.Selection.Range, 4, 4, ref defaultTableBehavior,
                    ref autoFitBehavior);
                    //Меняем стили созданных таблиц
                    Object style = "Классическая таблица 1";
                    wordtable1.set_Style(ref style);
                    style = "Сетка таблицы 3";
                    Object applystyle = true;
                    wordtable2.set_Style(ref style);
                    wordtable2.ApplyStyleFirstColumn = true;
                    wordtable2.ApplyStyleHeadingRows = true;
                    wordtable2.ApplyStyleLastRow = false;
                    wordtable2.ApplyStyleLastColumn = false;
                    break;
                case 2:
                    Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges;
                    Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
                    Object routeDocument = Type.Missing;
                    wordapp.Quit(ref saveChanges,
                                 ref originalFormat, ref routeDocument);
                    wordapp = null;
                    break;
               
                default:
                    Close();
                    break;
            }
        }
    }
}
