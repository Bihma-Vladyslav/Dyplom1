using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Word = Microsoft.Office.Interop.Word;

namespace Dyplom1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            db = new DBManager();
            db.connectTo();
        }
        DBManager db;
        int index = 0;
        string ind = "";
        Word.Application word;
        Word.Document doc;
        Word.Range r;
        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            db.selectall("Structure_Academic_Discipline", dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count != e.RowIndex + 1)
                try
                {
                    textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                    textBox5.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                    textBox6.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
                    textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
                    textBox8.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
                    textBox9.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
                    textBox10.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
                    textBox11.Text = dataGridView1.Rows[e.RowIndex].Cells[10].Value.ToString();
                    ind = db.getindex("Structure_Academic_Discipline", textBox1.Text, textBox2.Text);
                }
                catch
                {

                }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            String[] fields = { "Num_Section", "Num_Class", "Name_Section", "Total_Hours", "Lecture_Hours", "Workshop_Hours",
             "Practical_Hours","Laboratory_Hours", "IndepWorkStud_Hours", "Recommended_Books", "Forms_Means_Con" };
            String[] values = { textBox1.Text, textBox2.Text, "'" + textBox3.Text + "'", textBox4.Text,
            textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text,
             "'" + textBox10.Text + "'", "'" + textBox11.Text + "'"};

            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text)
                && !string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrEmpty(textBox4.Text))
            {
                if (!string.IsNullOrEmpty(textBox4.Text) && string.IsNullOrEmpty(textBox5.Text) && string.IsNullOrEmpty(textBox6.Text)
                    && string.IsNullOrEmpty(textBox7.Text) && string.IsNullOrEmpty(textBox8.Text) && string.IsNullOrEmpty(textBox9.Text))
                {
                    MessageBox.Show("Помилка! Будь ласка, обов'язково заповніть хоча б одне поле з кількістю годин, такі як: " +
                        "Лекційні години, Семінарські години, Практичні години, Лабораторні години, С.р.с години"); 
                }
                else
                { 
                db.insert("Structure_Academic_Discipline", fields, values);
                db.selectall("Structure_Academic_Discipline", dataGridView1);
                }
            }
            else
            {
                MessageBox.Show("Помилка! Будь ласка, обов'язково заповніть перші чотири поля: " +
                  "№ Розділу, № Заняття, Назва розділу та Усього годин");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
            { 
            db.delete("Structure_Academic_Discipline", "Num_Class=" + textBox2.Text);
            db.selectall("Structure_Academic_Discipline", dataGridView1);
            }
            else
            {
                MessageBox.Show("Помилка! Будь ласка, обов'язково заповніть друге поле: " +
                                   "№ Заняття, за яким відбувається функція видалення");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            String[] fields = { "Num_Section", "Num_Class", "Name_Section", "Total_Hours", "Lecture_Hours", "Workshop_Hours",
             "Practical_Hours","Laboratory_Hours", "IndepWorkStud_Hours", "Recommended_Books", "Forms_Means_Con" };
            String[] values = { textBox1.Text, textBox2.Text, "'" + textBox3.Text + "'", textBox4.Text,
            textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text,
             "'" + textBox10.Text + "'", "'" + textBox11.Text + "'"};
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text)
               && !string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrEmpty(textBox4.Text))
            {
                if (!string.IsNullOrEmpty(textBox4.Text) && string.IsNullOrEmpty(textBox5.Text) && string.IsNullOrEmpty(textBox6.Text)
                    && string.IsNullOrEmpty(textBox7.Text) && string.IsNullOrEmpty(textBox8.Text) && string.IsNullOrEmpty(textBox9.Text))
                {
                    MessageBox.Show("Помилка! Будь ласка, обов'язково заповніть хоча б одне поле з кількістю годин, такі як: " +
                        "Лекційні години, Семінарські години, Практичні години, Лабораторні години, С.р.с години");
                }
                else { 
                    db.update("Structure_Academic_Discipline", fields, values, "\"Index\"", ind);
                    db.selectall("Structure_Academic_Discipline", dataGridView1);
                    }
            }
            else
            {
                MessageBox.Show("Помилка! Будь ласка, обов'язково заповніть перші чотири поля: " +
                  "№ Розділу, № Заняття, Назва розділу та Усього годин");
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            index = 1;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            index = 2;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            index = 3;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            index = 4;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            index = 5;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            index = 6;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string[] sumResult = db.sumall("Structure_Academic_Discipline", fields, values);

            listBox3.Items.Add("");
            listBox3.Items.Add(sumResult[0]);
            listBox3.Items.Add(sumResult[1]);
            listBox3.Items.Add(sumResult[2]);
            listBox3.Items.Add(sumResult[3]);
            listBox3.Items.Add(sumResult[4]);
            listBox3.Items.Add(sumResult[5]);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //WORD DOCUMENT
        //---------------------------------------------------------------------------------------------------------------------
        private void button8_Click_1(object sender, EventArgs e)
        {
            try
            {
                word = new Word.Application();
                word.Visible = true;
                doc = word.Documents.Add();
                Word.Selection currentSelection = word.Application.Selection;
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.PageSetup.LeftMargin = word.CentimetersToPoints(1.5f);
                currentSelection.TypeText(label1.Text);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(label2.Text);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(label3.Text);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(label4.Text);
                currentSelection.TypeParagraph();
                int cur_pos = label1.Text.Length + label2.Text.Length + label3.Text.Length + label4.Text.Length;

                r = doc.Range(0, cur_pos + 4);
                // r.Bold = 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos + 4, cur_pos + 6);
                cur_pos = cur_pos + 6;
                string s1 = "Затверджую";
                string s2 = "Заступник директора";
                string s3 = "з навчально-методичної роботи";
                string s4 = "_____________ Анатолій МАЙДАН";
                string s5 = "«____» _________  2022 р.";

                currentSelection.TypeText(s1);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s2);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s3);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s4);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s5);

                r = doc.Range(cur_pos, cur_pos + s1.Length + s2.Length + s3.Length + s4.Length + s5.Length + 4);
                cur_pos = cur_pos + s1.Length + s2.Length + s3.Length + s4.Length + s5.Length + 4;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphRight;
                word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + 2);
                cur_pos = cur_pos + 2;

                string s6 = "_________________________";

                currentSelection.TypeText(s6);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s6.Length + 1);
                cur_pos = cur_pos + s6.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 16;
                r.Bold = 1;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                string s7 = "РОБОЧА НАВЧАЛЬНА ПРОГРАМА";

                currentSelection.TypeText(s7);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s7.Length + 2);
                cur_pos = cur_pos + s7.Length + 2;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.Bold = 1;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                string s8 = "галузь знань  12 «Інформаційні технології»";
                string s9 = "спеціальність 121  «Інженерія програмного забезпечення»";
                string s10 = "освітньо-професійна програма  «Інженерія програмного забезпечення»";
                string s11 = "освітньо-кваліфікаційний рівень            ";
                string s12 = "молодший спеціаліст";

                currentSelection.TypeText(s8);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s9);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s10);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s11);
                currentSelection.TypeText(s12);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s8.Length + s9.Length + s10.Length + s11.Length + s12.Length + 12);
                cur_pos = cur_pos + s8.Length + s9.Length + s10.Length + s11.Length + s12.Length + 12;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;

                string s13 = "20__/ 20__ навчальний рік";
                currentSelection.TypeText(s13);

                r = doc.Range(587, 612);
                cur_pos = 587 + 25;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.InsertBreak(Word.WdBreakType.wdPageBreak);


                string s14 = "Робоча програма навчальної дисципліни ";
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1.15f);
                string s15 = "_______________________________________ ";

                currentSelection.TypeText(s14);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s15);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s14.Length + s15.Length + 3);
                cur_pos = cur_pos + s14.Length + s15.Length + 3;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;

                string s16 = "     (назва навчальної дисципліни)";

                currentSelection.TypeText(s16);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s16.Length + 2);
                cur_pos = cur_pos + s16.Length + 2;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;


                string s17 = "для здобувачів фахової передвищої освіти за спеціальністю 121  «Інженерія програмного забезпечення», " +
                    "освітньо-професійною програмою «Інженерія програмного забезпечення». ";

                currentSelection.TypeText(s17);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s17.Length + 3);
                cur_pos = cur_pos + s17.Length + 3;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;

                string s18 = "Розробники: Круш Ольга спеціаліст вищої категорії, викладач-методист, викладач спецдисциплін";

                currentSelection.TypeText(s18);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s18.Length + 5);
                cur_pos = cur_pos + s18.Length + 5;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;


                string s19 = "Робочу програму схвалено на засіданні циклової випускової комісії спеціальності " +
                    "121 «Інженерія програмного забезпечення»";
                currentSelection.TypeText(s19);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s19.Length + 3);
                cur_pos = cur_pos + s19.Length + 3;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;

                string s20 = "Протокол від «__» ______ ____ року № _";
                currentSelection.TypeText(s20);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s20.Length + 2);
                cur_pos = cur_pos + s20.Length + 2;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;

                string s21 = "Голова циклової випускової комісії спеціальності 121 «Інженерія програмного забезпечення»";
                currentSelection.TypeText(s21);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s21.Length + 1);
                cur_pos = cur_pos + s21.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphJustify;

                string s22 = "                   _________________________ (____";
                currentSelection.TypeText(s22);

                r = doc.Range(cur_pos, cur_pos + s22.Length);
                cur_pos = cur_pos + s22.Length;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphJustify;

                string s23 = "Олена ВИСОЦЬКА_";
                currentSelection.TypeText(s23);

                r = doc.Range(cur_pos, cur_pos + s23.Length);
                cur_pos = cur_pos + s23.Length;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.Underline = Word.WdUnderline.wdUnderlineSingle;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphJustify;

                string s24 = "___)";
                currentSelection.TypeText(s24);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s24.Length + 1);
                cur_pos = cur_pos + s24.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphJustify;

                string s25 = "                                                                 (підпис)" +
                   "                                                   (прізвище та ініціали)";
                currentSelection.TypeText(s25);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + s25.Length + 2);
                cur_pos = cur_pos + s25.Length + 2;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 9;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;

                currentSelection.InsertBreak(Word.WdBreakType.wdPageBreak);
                string s26 = "1. Опис навчальної дисципліни";
                currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0.62f);
                currentSelection.TypeText(s26);
                currentSelection.TypeParagraph();
                currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0f);

                r = doc.Range(cur_pos, cur_pos + s26.Length + 4);
                cur_pos = cur_pos + s26.Length + 4;
                // r.ListFormat.ApplyNumberDefault();//ось тут починається з двох
                r.Bold = 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;


                //THERE IS TABLE
                //-------------------------------------------------
                //Selection.SelectCell
                //Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                //Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter

                r = doc.Range(cur_pos, cur_pos); //баг якщо змінювати розміщення тексту 2022/2023 рік
                cur_pos = cur_pos;



                Word.Table t = doc.Tables.Add(r, 3, 3);
                t.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                t.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.TypeText("Найменування показників");
                currentSelection.MoveRight();

                r = t.Cell(1, 1).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.TypeText("Спеціальність, ОПП, освітньо-професійний ступінь освіти");
                currentSelection.MoveRight();

                r = t.Cell(1, 2).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.TypeText("Характеристика навчальної дисципліни");
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: false);

                r = t.Cell(1, 3).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.MoveRight();
                currentSelection.MoveRight();
                currentSelection.TypeText("денна форма навчання");
                currentSelection.MoveRight();
                currentSelection.MoveRight();

                r = t.Cell(2, 3).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;


                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.TypeText("Кількість кредитів – ");
                currentSelection.MoveRight();

                r = t.Cell(3, 1).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                string tab_s_1 = "Спеціальність:";
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.TypeText(tab_s_1);
                currentSelection.TypeParagraph();

                r = t.Cell(3, 2).Range;
                r = doc.Range(cur_pos + 164, cur_pos + 164 + tab_s_1.Length + 1);
                cur_pos = cur_pos + 164 + tab_s_1.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.Bold = 1;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                string tab_s_2 = "121  «Інженерія програмного забезпечення»";
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);

             currentSelection.TypeText(tab_s_2);
             currentSelection.MoveRight();
             r = doc.Range(cur_pos, cur_pos + tab_s_2.Length + 2);
             cur_pos = cur_pos + tab_s_2.Length + 2;
             r.Font.Name = " Times New Roman ";
             r.Bold = 0;
             r.Font.Size = 12;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphCenter;
            
             currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
             currentSelection.SelectCell();
             currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
             currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
             currentSelection.TypeText("За вибором");
             currentSelection.MoveRight();
             currentSelection.MoveRight();
                
             r = t.Cell(3, 3).Range;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 12;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphCenter;


            currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
            currentSelection.SelectCell();
            currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            currentSelection.TypeText("Розділів (змістових модулів) –  ");
            currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: false);
            r = t.Cell(4, 1).Range;
            r.Font.Name = " Times New Roman ";
            r.Font.Size = 12;
            r.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.MoveRight();
                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.TypeText("Загальна кількість годин –  ");
                currentSelection.MoveRight();
                currentSelection.MoveUp();
                currentSelection.MoveLeft();

                r = t.Cell(5, 1).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                string tab_s_3 = "освітньо-професійна програма: ";
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.TypeText(tab_s_3);
                currentSelection.TypeParagraph();

                r = t.Cell(4, 2).Range;
                r = doc.Range(cur_pos + 44, cur_pos + 44 + tab_s_3.Length + 1);
                cur_pos = cur_pos + 44 + tab_s_3.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.Bold = 1;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
                
                string tab_s_4 = "«Інженерія програмного забезпечення»";
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);

                currentSelection.TypeText(tab_s_4);
                currentSelection.MoveRight();
                r = doc.Range(cur_pos, cur_pos + tab_s_4.Length); //1793+30+1 = 1824, 1793+30+1+36+1 = 1861 
                cur_pos = cur_pos + tab_s_4.Length;
                r.Font.Name = " Times New Roman ";
                r.Bold = 0;
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.TypeText("Рік підготовки");
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: false);

                r = t.Cell(4, 3).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.MoveRight();
                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.TypeText("20__-й");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("20__-й");
                //
                currentSelection.SelectCell();

                r = t.Cell(5, 3).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);
                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveUp();
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);

                currentSelection.MoveDown();
                currentSelection.TypeText("Семестр");
                currentSelection.SelectCell();

                r = t.Cell(6, 3).Range;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 12;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);
                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();

                currentSelection.TypeText("_-й");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("_-й");
                currentSelection.SelectCell();

                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);
                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.MoveUp();
                currentSelection.SelectCell();

                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);
                currentSelection.MoveDown();
                currentSelection.TypeText("Лекції");

                currentSelection.InsertRowsBelow(1);
                currentSelection.MoveLeft();
                currentSelection.TypeText("Тижневих годин для денної форми навчання:");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("аудиторних: ");
                currentSelection.TypeParagraph();

                currentSelection.TypeText("1 семестр – __ год");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("2 семестр – __ год");
                currentSelection.TypeParagraph();

                currentSelection.TypeText("самостійної роботи:");
                currentSelection.TypeParagraph();

                currentSelection.TypeText("1 семестр – __ год");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("2 семестр – __ год");
                currentSelection.TypeParagraph();
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                string tab_s_5 = "Освітньо-";
                currentSelection.MoveRight();
                currentSelection.TypeText(tab_s_5);

                string tab_s_6 = "кваліфікаційний рівень:";

                currentSelection.TypeParagraph();
                //  r = t.Cell(6, 2).Range;
                r = doc.Range(cur_pos + 246, cur_pos + 246 + tab_s_5.Length + 1); //1793+30+1 = 1824, 1793+30+1+36+1 = 1861 
                cur_pos = cur_pos + 246 + tab_s_5.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Bold = 0;
                r.Font.Size = 12;
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                currentSelection.TypeText(tab_s_6);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + tab_s_6.Length + 1); //1793+30+1 = 1824, 1793+30+1+36+1 = 1861 
                cur_pos = cur_pos + tab_s_6.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Bold = 0;
                r.Font.Size = 12;

                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                string tab_s_7 = "молодший спеціаліст";


                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.TypeText(tab_s_7);
                currentSelection.TypeParagraph();

                r = doc.Range(cur_pos, cur_pos + tab_s_7.Length + 1); //1793+30+1 = 1824, 1793+30+1+36+1 = 1861 
                cur_pos = cur_pos + tab_s_7.Length + 1;
                r.Font.Name = " Times New Roman ";
                r.Bold = 1;
                r.Font.Size = 12;

                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                currentSelection.MoveRight();
                currentSelection.TypeText("_ год.");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("_ год.");
                currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1f);
                currentSelection.SelectCell();
                currentSelection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                currentSelection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);

                currentSelection.MoveRight();
                currentSelection.MoveDown();
                currentSelection.TypeText("Практичні, семінарські");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.TypeText("_ год.");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("_ год.");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.MoveUp();
                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.TypeText("Лабораторні");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.TypeText("_ год.");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("_ год.");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.MoveUp();
                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.TypeText("Самостійна робота");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.TypeText("_ год.");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("_ год.");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.MoveUp();
                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);

                currentSelection.MoveDown();
                currentSelection.TypeText("Індивідуальні завдання: ");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.TypeText("_ год.");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveLeft();
                currentSelection.MoveDown();
                currentSelection.MoveLeft();
                currentSelection.MoveLeft();
                currentSelection.Cells.Split(NumRows: 2, NumColumns: 1, MergeBeforeSplit: true);

                currentSelection.MoveDown();
                currentSelection.MoveUp();
                currentSelection.MoveUp();
                currentSelection.TypeText("Вид контролю: ");

                currentSelection.MoveDown();
                currentSelection.TypeText("залік.");
                currentSelection.TypeParagraph();
                currentSelection.TypeText("екзамен");
                currentSelection.SelectCell();
                currentSelection.Cells.Split(NumRows: 1, NumColumns: 2, MergeBeforeSplit: true);

                currentSelection.MoveDown();
                currentSelection.TypeParagraph();

             //TABLES ENDS
             //-------------------------------
             currentSelection.ParagraphFormat.LineSpacing = word.LinesToPoints(1.15f);

             string s27 = "Примітка.";
             currentSelection.TypeText(s27);
             currentSelection.TypeParagraph();

             r = doc.Range(2340, 2340 + s27.Length + 1);
             cur_pos = 2340 + s27.Length + 1;
             r.Bold = 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s28 = "Співвідношення кількості годин аудиторних занять " +
                 "до самостійної роботи здобувача фахової передвищої освіти становить:";
             currentSelection.TypeText(s28);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s28.Length + 1);
             cur_pos = cur_pos + s28.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s29 = "Для денної форми навчання - __ %.";
             currentSelection.TypeText(s29);
             currentSelection.TypeParagraph();
             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0.62f);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s29.Length + 2);
             cur_pos = cur_pos + s29.Length + 2;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s30 = "2. Мета та завдання навчальної дисципліни";
             currentSelection.TypeText(s30);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s30.Length + 1);
             cur_pos = cur_pos + s30.Length + 1;
            // r.ListFormat.ApplyNumberDefault();
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.Bold = 1;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s31 = "Мета ";
             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0f);
             currentSelection.TypeText(s31);

             r = doc.Range(cur_pos, cur_pos + s31.Length);
             cur_pos = cur_pos + s31.Length;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.Bold = 1;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s32 = "дисципліни «(Назва дисципліни)» є ...";
             currentSelection.TypeText(s32);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s32.Length + 1);
             cur_pos = cur_pos + s32.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s33 = "Завдання ";
             currentSelection.TypeText(s33);
             //currentSelection.PageSetup.LeftMargin = word.CentimetersToPoints(2.5f);
             currentSelection.ParagraphFormat.FirstLineIndent = word.CentimetersToPoints(1.25f);
             r = doc.Range(cur_pos, cur_pos + s33.Length);
             cur_pos = cur_pos + s33.Length;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.Bold = 1;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s34 = "дисципліни «(Назва дисципліни)» полягає ...";
             currentSelection.TypeText(s34);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s34.Length + 1);
             cur_pos = cur_pos + s34.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             //currentSelection.ParagraphFormat.LineSpacing = word.CentimetersToPoints(1f);

             string s35 = "Головною задачею дисципліни є:";
             currentSelection.TypeText(s35);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s35.Length + 1);
             cur_pos = cur_pos + s35.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s36 = "- (текст);";
             currentSelection.TypeText(s36);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s36.Length + 1);
             cur_pos = cur_pos + s36.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s37 = "- (текст).";
             currentSelection.TypeText(s37);
             currentSelection.TypeParagraph();
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s37.Length + 2);
             cur_pos = cur_pos + s37.Length + 2;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;


             string s38 = "У результаті вивчення навчальної дисципліни здобувачі фахової передвищої освіти повинні:";
             currentSelection.TypeText(s38);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s38.Length + 1);
             cur_pos = cur_pos + s38.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s39 = "Знати";
             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0.60f);
             currentSelection.ParagraphFormat.FirstLineIndent = word.CentimetersToPoints(0f);
             currentSelection.TypeText(s39);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s39.Length + 1);
             cur_pos = cur_pos + s39.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(2f);
             currentSelection.ParagraphFormat.FirstLineIndent = word.CentimetersToPoints(-0.67f);
             string s40 = "* (текст);";
             currentSelection.TypeText(s40);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s40.Length + 1);
             cur_pos = cur_pos + s40.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s41 = "* (текст).";
             currentSelection.TypeText(s41);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s41.Length + 1);
             cur_pos = cur_pos + s41.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;


             string s42 = "Вміти";
             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0.60f);
             currentSelection.ParagraphFormat.FirstLineIndent = word.CentimetersToPoints(0f);
             currentSelection.TypeText(s42);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s42.Length + 1);
             cur_pos = cur_pos + s42.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s43 = "* (текст);";
             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(2f);
             currentSelection.ParagraphFormat.FirstLineIndent = word.CentimetersToPoints(-0.67f);
             currentSelection.TypeText(s43);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s43.Length + 1);
             cur_pos = cur_pos + s43.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s44 = "* (текст).";
             currentSelection.TypeText(s44);
             currentSelection.TypeParagraph();
             currentSelection.ParagraphFormat.LeftIndent = word.CentimetersToPoints(0f);
             currentSelection.ParagraphFormat.FirstLineIndent = word.CentimetersToPoints(1.5f);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s44.Length + 2);
             cur_pos = cur_pos + s44.Length + 2;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;

             string s45= "Сформовані компетентності";
             currentSelection.TypeText(s45);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s45.Length + 1);
             cur_pos = cur_pos + s45.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.Bold = 1;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s46 = "Загальні компетентності: ";
             currentSelection.TypeText(s46);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s46.Length + 1);
             cur_pos = cur_pos + s46.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Italic = 1;
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s47 = "ЗК05. Знання та розуміння предметної області та розуміння професійної діяльності.";
             currentSelection.TypeText(s47);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s47.Length + 1);
             cur_pos = cur_pos + s47.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s48 = "ЗК07. Здатність застосовувати знання у практичних ситуаціях. ";
             currentSelection.TypeText(s48);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s48.Length + 1);
             cur_pos = cur_pos + s48.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s49 = "Фахові компетентності: ";
             currentSelection.TypeText(s49);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s49.Length + 1);
             cur_pos = cur_pos + s49.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Italic = 1;
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s50 = "(текст).";
             currentSelection.TypeText(s50);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s50.Length + 1);
             cur_pos = cur_pos + s50.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             currentSelection.TypeText(s50);
             currentSelection.TypeParagraph();
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s50.Length + 2);
             cur_pos = cur_pos + s50.Length + 2;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s51 = "Залежно від типу обчислювальної техніки, складу наявного науково-методичного та програмного забезпечення " +
                 "викладач може самостійно добирати методичні шляхи розв’язування освітніх завдань дисципліни, " +
                 "вносити необхідні корективи в порядок вивчення тем програми, а також змінювати кількість годин, " +
                 "необхідних для засвоєння навчального матеріалу з окремих тем програми. " +
                 "Окремі питання програми можуть вивчатися тільки в порядку ознайомлення.";
             currentSelection.TypeText(s51);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s51.Length + 1);
             cur_pos = cur_pos + s51.Length + 1;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;

             string s52 = "Матеріал курсу є базовим при вивченні здобувачами фахової передвищої освіти дисциплін учбового плану, " +
                 "пов'язаних із створенням різноманітних комп`ютерних систем. Отримані знання дозволять здобувачам фахової передвищої освіти використовувати " +
                 "методи інформаційного моделювання при вивченні інших інженерних дисциплін, виконанні курсових і дипломних робіт.";
             currentSelection.TypeText(s52);
             currentSelection.TypeParagraph();
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s52.Length + 2);
             cur_pos = cur_pos + s52.Length + 2;
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;
             
             string s53 = "3. Програма навчальної дисципліни";
             currentSelection.TypeText(s53);
             currentSelection.TypeParagraph();

             r = doc.Range(cur_pos, cur_pos + s53.Length + 1);
             cur_pos = cur_pos + s53.Length + 1;
            // r.ListFormat.ApplyNumberDefault();
             r.Font.Name = " Times New Roman ";
             r.Font.Size = 14;
             r.Bold = 1;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;
                
            string s54 = "1. Розділ (змістовий модуль) 1. Моделі конструювання";
            currentSelection.TypeText(s54);
            currentSelection.TypeParagraph();
            string s55 = "2. Розділ (змістовий модуль) 2. Планування конструювання";
            currentSelection.TypeText(s55);
            currentSelection.TypeParagraph();
            string s56 = "3. Розділ (змістовий модуль) 3. Мови конструювання";
            currentSelection.TypeText(s56);
            currentSelection.TypeParagraph();
            string s57 = "4. Розділ (змістовий модуль) 4. Інтеграція";
            currentSelection.TypeText(s57);
            currentSelection.TypeParagraph();
            string s58 = "5. Розділ (змістовий модуль) 5. Якість конструювання";
            currentSelection.TypeText(s58);
            currentSelection.TypeParagraph();
            string s59 = "6. Розділ (змістовий модуль) 6. Шаблони проектування";
            currentSelection.TypeText(s59);
            currentSelection.TypeParagraph();
            currentSelection.TypeParagraph();

            r = doc.Range(cur_pos, cur_pos + s54.Length + s55.Length + s56.Length + s57.Length + s58.Length + s59.Length + 7);
            cur_pos = cur_pos + s54.Length + s55.Length + s56.Length + s57.Length + s58.Length + s59.Length + 7;
           // r.ListFormat.ApplyNumberDefault();
            r.Font.Name = " Times New Roman ";
            r.Font.Size = 14;
            r.Bold = 1;
            r.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphLeft;

            string s60 = "4. Структура навчальної дисципліни";
            currentSelection.TypeText(s60);
            currentSelection.TypeParagraph();

            r = doc.Range(cur_pos, cur_pos + s60.Length + 1);
            // r = doc.Range(3924, 3924 + s60.Length + 1); //помилка з діапазоном
            cur_pos = cur_pos + s60.Length + 1;
            // r.ListFormat.ApplyNumberDefault();
        // r = doc.Range(3594, 3630);
        // cur_pos = 3594 + 36;
            r.Font.Name = " Times New Roman ";
            r.Font.Size = 14;
            r.Bold = 1;
            r.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphLeft;
                             
            
             r = doc.Range(cur_pos, cur_pos);

             Word.Table t1 = doc.Tables.Add(r, dataGridView1.RowCount, dataGridView1.ColumnCount);
             t1.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
             t1.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                //currentSelection.Tables(1).Columns(1).SetWidth ColumnWidth:= 21.05, RulerStyle:= _wdAdjustNone //потрібно випробувати
                //currentSelection.Tables(1).Columns(2).SetWidth ColumnWidth:= 42.5, RulerStyle:= _wdAdjustNone //потрібно випробувати

                for(int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    currentSelection.TypeText(dataGridView1.Columns[j].HeaderText);
                    currentSelection.MoveRight();
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                { 
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        if(dataGridView1.Rows[i].Cells[j].Value != null)
                        { 
                            currentSelection.TypeText(dataGridView1.Rows[i].Cells[j].Value.ToString());
                        }
                        currentSelection.MoveRight();
                    }
                    currentSelection.TypeParagraph();
                }
                t1.Columns[1].Width = 28f;
                t1.Columns[2].Width = 28f;
                t1.Columns[3].Width = 130f;
                t1.Columns[4].Width = 28f;
                t1.Columns[5].Width = 28f;
                t1.Columns[6].Width = 28f;
                t1.Columns[7].Width = 28f;
                t1.Columns[8].Width = 28f;
                t1.Columns[9].Width = 28f;
                t1.Columns[10].Width = 80f;
                t1.Columns[11].Width = 90f;

                t1.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 1).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 2).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                t1.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                t1.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 4).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 5).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 5).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 6).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 6).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 7).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 7).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 8).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 8).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 8).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 9).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                t1.Cell(1, 9).Range.Orientation = Word.WdTextOrientation.wdTextOrientationUpward;
                t1.Cell(1, 9).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                t1.Cell(1, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                t1.Cell(1, 10).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                t1.Cell(1, 11).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                t1.Cell(1, 11).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                t1.Cell(1, 3).Range.Text = dataGridView1.Columns[2].HeaderText + " (змістових модулів) і тем";
                currentSelection.MoveRight();

                word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
                String[] values = new string[7];

                for (int j = 1; j <= 6; j++)
                {

                    string condition = "Num_Section = " + j;
                    string[] sumResult = db.sum("Structure_Academic_Discipline", fields, values, condition);


                    currentSelection.TypeText("Разом за розділом(змістовим модулем)" + j.ToString() + ":");
                    currentSelection.TypeParagraph();

                    for (int i = 1; i < listBox2.Items.Count; i++)
                    {
                        currentSelection.TypeText(listBox2.Items[i].ToString());
                        if (string.IsNullOrEmpty(sumResult[i - 1]))
                        {
                            currentSelection.TypeText("0");
                        }
                        else
                        {
                            currentSelection.TypeText(sumResult[i - 1]);
                        }
                        currentSelection.TypeParagraph();
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                word.Quit();
            }
            finally
            {
                // word.Quit();
                word = null;
                doc = null;
            }
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}