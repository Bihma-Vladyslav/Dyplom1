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

        private void button8_Click_1(object sender, EventArgs e)
        {
            try
            {
                //for
                word = new Word.Application();
                word.Visible = true;
                doc = word.Documents.Add();
                Word.Selection currentSelection = word.Application.Selection;
                currentSelection.TypeText(label1.Text);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(label2.Text);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(label3.Text);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(label4.Text);
                currentSelection.TypeParagraph();
                int cur_pos = label1.Text.Length + label2.Text.Length + label3.Text.Length + label4.Text.Length;
                r = doc.Range(0, cur_pos + 1);
                // r.Bold = 1;
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;

                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();

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
                r = doc.Range(cur_pos + 1, cur_pos + s1.Length + s2.Length + s3.Length + s4.Length + s5.Length + 1);
                r.Font.Name = " Times New Roman ";
                r.Font.Size = 14;
                r.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphRight;
                word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                currentSelection.TypeParagraph();

                string s6 = "Конструювання програмного забезпечення";
                string s7 = "РОБОЧА НАВЧАЛЬНА ПРОГРАМА";
                string s8 = "галузь знань  12 «Інформаційні технології»";
                string s9 = "спеціальність 121  «Інженерія програмного забезпечення»";
                string s10 = "освітньо-професійна програма  «Інженерія програмного забезпечення»";
                string s11 = "освітньо-кваліфікаційний рівень            ";
                string s12 = "молодший спеціаліст";
                string s13 = "2022/ 2023 навчальний рік";


                currentSelection.TypeText(s6);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s7);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s8);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s9);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s10);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s11);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s12);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s13);
                

                //  r = doc.Range(cur_pos + 1, cur_pos + s6.Length + s7.Length + s8.Length + s9.Length + s10.Length + s11.Length + s12.Length + s13.Length);
                // r.Bold = 1;
                r.Font.Name = " Times New Roman ";
                 r.Font.Size = 14;
                currentSelection.InsertBreak(Word.WdBreakType.wdPageBreak);
                // r.ParagraphFormat.Alignment =
                //  Word.WdParagraphAlignment.wdAlignParagraphCenter;
                string s14 = "Робоча програма навчальної дисципліни ";
                string s15 = "__Конструювання програмного забезпечення__ ";
                string s16 = "(     назва навчальної дисципліни)";
                string s17 = "для здобувачів фахової передвищої освіти за спеціальністю 121  «Інженерія програмного забезпечення», " +
                    "освітньо-професійною програмою «Інженерія програмного забезпечення». ";
                string s18 = "Розробники: Круш Ольга спеціаліст вищої категорії, викладач-методист, викладач спецдисциплін";
                string s19 = "Робочу програму схвалено на засіданні циклової випускової комісії спеціальності " +
                    "121 «Інженерія програмного забезпечення»";
                string s20 = "Протокол від «26» серпня 2022 року № 1";
                string s21 = "Голова циклової випускової комісії спеціальності 121 «Інженерія програмного забезпечення»";
                string s22 = "                   _________________________ (____Олена ВИСОЦЬКА___)";
                string s23 = "                                                                 (підпис)" +
                    "                                                   (прізвище та ініціали)";

                currentSelection.TypeText(s14);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s15);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s16);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s17);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s18);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s19);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s20);
                currentSelection.TypeParagraph();
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s21);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s22);
                currentSelection.TypeParagraph();
                currentSelection.TypeText(s23);
                currentSelection.InsertBreak(Word.WdBreakType.wdPageBreak);

                string s24 = "1.	Опис навчальної дисципліни";
                currentSelection.TypeParagraph();
                r = doc.Range(cur_pos+1, cur_pos+1);

                Word.Table t = doc.Tables.Add(r, 4, 3);
                t.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                t.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                currentSelection.TypeText("TEST");

                r = t.Cell(1, 1).Range;
                for (int row = 1; row <= 4; row++)
                {
                    for (int column = 1; column <= 3; column++)
                    {
                        string cellText = $"Найменування показників {row}, Спеціальність, ОПП, освітньо-професійний ступінь освіти {column}";
                        t.Cell(row, column).Range.Text = cellText;
                    }
                }
               // string s25 = "Найменування показників";
               // currentSelection.TypeText(s25);
                //string s26 = "Спеціальність, ОПП, освітньо-професійний ступінь освіти";
                //currentSelection.TypeText(s26);
            }


            /* r = doc.Range(cur_pos + 1, cur_pos + textBox2.Text.Length);
             r.Italic = 1;
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphRight;
             currentSelection.TypeParagraph();
             cur_pos = cur_pos + textBox2.Text.Length + 1;
             r = doc.Range(cur_pos, cur_pos);
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphJustify;
             r.Font.Name = " Times New Roman " ;
             currentSelection.TypeText(listBox1.Text);
             currentSelection.TypeParagraph();
             currentSelection.TypeText(label2.Text + " " +textBox3.Text);
             cur_pos = cur_pos + listBox1.Text.Length + 1;
             r = doc.Range(cur_pos, cur_pos + label2.Text.Length +
             textBox3.Text.Length + 1);
             r.ParagraphFormat.Alignment =
             Word.WdParagraphAlignment.wdAlignParagraphLeft;
             r.Underline = Word.WdUnderline.wdUnderlineDotted;
             r.Font.Name = " Times New Roman " ;
             cur_pos = cur_pos + label2.Text.Length + textBox3.Text.Length + 1;
             r = doc.Range(cur_pos + 1, cur_pos + 1);
             Word.Table t = doc.Tables.Add(r, dataGridView1.RowCount,
             dataGridView1.ColumnCount);
             t.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
             t.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;}*/
            /*for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                currentSelection.TypeText(dataGridView1.Columns[j].HeaderText);
                currentSelection.MoveRight();
            }
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        currentSelection.TypeText(dataGridView1.Rows[i].Cells[j].Value.ToString());
                    }
                    currentSelection.MoveRight();
                }
            currentSelection.MoveRight();
            r = t.Cell(1, 1).Range;
            r.Bold = 1;
            r.Font.Color = Word.WdColor.wdColorBlue;
            currentSelection.TypeText(" end of table ");
            word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            word.Documents.Save(false);
        }*/
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