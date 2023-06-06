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
        int index = 0;
        string ind = "";
        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            db.selectall1("Topics_Seminar_Classes", dataGridView1);
            button14.Tag = 1;
            button15.Tag = 0;
            button16.Tag = 0;
            button17.Tag = 0;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            if (button14.Tag.Equals(1))
            {
                label8.Text = "№ з/п";

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            db.selectall1("Topics_Practical_Classes", dataGridView1);
            button14.Tag = 0;
            button15.Tag = 1;
            button16.Tag = 0;
            button17.Tag = 0;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            if (button15.Tag.Equals(1))
            {
                label8.Text = "№ з/п";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            db.selectall2("Topics_Laboratory_Works", dataGridView1);
            button14.Tag = 0;
            button15.Tag = 0;
            button16.Tag = 1;
            button17.Tag = 0;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            if (button16.Tag.Equals(1))
            {
                label8.Text = "пн";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            db.selectall2("Topics_Independent_Works", dataGridView1);
            button14.Tag = 0;
            button15.Tag = 0;
            button16.Tag = 0;
            button17.Tag = 1;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            if (button17.Tag.Equals(1))
            {
                label8.Text = "пн";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text };
            //треба прописати, за якою умовою додаються поля зі значеннями в яку конкретно таблицю
            //бо всі таблиці мають однакові поля, але як зрозуміти, в яку саме користувач хоче вписати дані цим insert-ом
            if (button14.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Topic_Name", "Number_Hours" };
                db.insert("Topics_Seminar_Classes", fields, values);
                db.selectall1("Topics_Seminar_Classes", dataGridView1);
            }
            if (button15.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Topic_Name", "Number_Hours" };
                db.insert("Topics_Practical_Classes", fields, values);
                db.selectall1("Topics_Practical_Classes", dataGridView1);
            }
            if (button16.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Name_Class", "Number_Hours" };
                db.insert("Topics_Laboratory_Works", fields, values);
                db.selectall2("Topics_Laboratory_Works", dataGridView1);
            }
            if (button17.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Name_Class", "Number_Hours" };
                db.insert("Topics_Independent_Works", fields, values);
                db.selectall2("Topics_Independent_Works", dataGridView1);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Rows.Count != e.RowIndex + 1)
                try
                {
                    //тут трабл якщо натискаєш на пусту клітинку
                    textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                    textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    ind = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                }
                catch
                {

                }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (button14.Tag.Equals(1))
            {
                db.delete("Topics_Seminar_Classes", "Number_Sequence=" + textBox1.Text);
                db.selectall1("Topics_Seminar_Classes", dataGridView1);
            }
            if (button15.Tag.Equals(1))
            {
                db.delete("Topics_Practical_Classes", "Number_Sequence=" + textBox1.Text);
                db.selectall1("Topics_Practical_Classes", dataGridView1);
            }
            if (button16.Tag.Equals(1))
            {
                db.delete("Topics_Laboratory_Works", "Number_Sequence=" + textBox1.Text);
                db.selectall2("Topics_Laboratory_Works", dataGridView1);
            }
            if (button17.Tag.Equals(1))
            {
                db.delete("Topics_Independent_Works", "Number_Sequence=" + textBox1.Text);
                db.selectall2("Topics_Independent_Works", dataGridView1);
            }
            //db.delete("Topics_Seminar_Classes", "Number_Sequence=" + textBox1.Text);
            //db.selectall("Topics_Seminar_Classes", dataGridView1);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (button14.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Topic_Name", "Number_Hours"};
                String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text };
                db.update("Topics_Seminar_Classes", fields, values, "\"Number_Sequence\"", ind);
                db.selectall1("Topics_Seminar_Classes", dataGridView1);
            }
            if (button15.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Topic_Name", "Number_Hours" };
                String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text };
                db.update("Topics_Practical_Classes", fields, values, "\"Number_Sequence\"", ind);
                db.selectall1("Topics_Practical_Classes", dataGridView1);
            }
            if (button16.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Name_Class", "Number_Hours" };
                String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text };
                db.update("Topics_Laboratory_Works", fields, values, "\"Number_Sequence\"", ind);
                db.selectall2("Topics_Laboratory_Works", dataGridView1);
            }
            if (button17.Tag.Equals(1))
            {
                String[] fields = { "Number_Sequence", "Name_Class", "Number_Hours" };
                String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text };
                db.update("Topics_Independent_Works", fields, values, "\"Number_Sequence\"", ind);
                db.selectall2("Topics_Independent_Works", dataGridView1);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            this.Close();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
    }
}