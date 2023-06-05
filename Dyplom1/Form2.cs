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
        }

        private void button2_Click(object sender, EventArgs e)
        {
            db.selectall1("Topics_Practical_Classes", dataGridView1);
            button14.Tag = 0;
            button15.Tag = 1;
            button16.Tag = 0;
            button17.Tag = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            db.selectall2("Topics_Laboratory_Classes", dataGridView1);
            button14.Tag = 0;
            button15.Tag = 0;
            button16.Tag = 1;
            button17.Tag = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            db.selectall3("Topics_Independent_Works", dataGridView1);
            button14.Tag = 1;
            button15.Tag = 0;
            button16.Tag = 0;
            button17.Tag = 17;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            String[] fields = { "Number_Suquence", "Topic_Name", "Number_Hours" };
            String[] values = { textBox1.Text, "'" + textBox2.Text + "'", textBox3.Text };
            //треба прописати, за якою умовою додаються поля зі значеннями в яку конкретно таблицю
            //бо всі таблиці мають однакові поля, але як зрозуміти, в яку саме користувач хоче вписати дані цим insert-ом
            if (button14.Tag.Equals(1))
            {
                db.insert("Topics_Seminar_Classes", fields, values);
                db.selectall("Topics_Seminar_Classes", dataGridView1);
            }
            if (button15.Tag.Equals(1))
            {
                db.insert("Topics_Practical_Classes", fields, values);
                db.selectall("Topics_Practical_Classes", dataGridView1);
            }
            if (button16.Tag.Equals(1))
            {
                db.insert("Topics_Laboratory_Classes", fields, values);
                db.selectall("Topics_Laboratory_Classes", dataGridView1);
            }
            if (button17.Tag.Equals(1))
            {
                db.insert("Topics_Independent_Works", fields, values);
                db.selectall("Topics_Independent_Works", dataGridView1);
            }
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
        
    }
}