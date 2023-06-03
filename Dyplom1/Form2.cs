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

namespace Dyplom1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            db = new DBManager();
            db.connectTo();
        }
        DBManager db;

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            db.selectall("Topics_Seminar_Classes", dataGridView1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            db.selectall("Topics_Practical_Classes", dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            db.selectall("Topics_Laboratory_Classes", dataGridView1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            db.selectall("Topics_Independent_Works", dataGridView1);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            String[] fields = { "Number_Suquence", "Topic_Name", "Number_Hours"};
            String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text};
            //треба прописати, за якою умовою додаються поля зі значеннями в яку конкретно таблицю
            //бо всі таблиці мають однакові поля, але як зрозуміти, в яку саме користувач хоче вписати дані цим insert-ом
            /* if()
             db.insert("Structure_Academic_Discipline", fields, values);
             db.selectall("Structure_Academic_Discipline", dataGridView1);*/
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //тут трабл якщо натискаєш на пусту клітинку
                textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            }
            catch
            {

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //треба прописати, за якою умовою видаляється рядок з полями і значеннями в них, в яку конкретно таблицю
            //бо всі таблиці мають однакові поля, але як зрозуміти, який саме рядок користувач хоче видалити цим delet-ом
            //if
            /* db.delete("Structure_Academic_Discipline", "Number_Suquence=" + textBox1.Text); //!!!!тут правильний тільки "Number_Suquence="
             db.selectall("Structure_Academic_Discipline", dataGridView1);*/
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //треба прописати, за якою умовою оновлюється поля  і значеннями в них, в яку конкретно таблицю
            //бо всі таблиці мають однакові поля, але як зрозуміти, яке саме поле/поля користувач хоче змінити цим update-ом
            /* String[] fields = { "Num_Section", "Num_Class", "Name_Section", "Total_Hours", "Lecture_Hours", "Workshop_Hours",
              "Practical_Hours","Laboratory_Hours", "IndepWorkStud_Hours", "Recommended_Books", "Forms_Means_Con" };
             String[] values = { textBox1.Text, textBox2.Text, "'" + textBox3.Text + "'", textBox4.Text,
             textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text,
              "'" + textBox10.Text + "'", "'" + textBox11.Text + "'"};

             db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text);
             db.selectall("Structure_Academic_Discipline", dataGridView1);*/
        }
    }
}