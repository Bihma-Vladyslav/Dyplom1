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

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            db = new DBManager();
            db.connectTo();
        }
        DBManager db;
        private void button7_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
          /*  var data = db.selectall("Structure_Academic_Discipline");
            dataGridView1.Columns.Clear();
            if (data != null)
            {
                for (int i = 0; i < data[0].Count; i++)
                {
                    dataGridView1.Columns.Add("col" + i.ToString(), "col" + i.ToString());
                }
                for (int i = 0; i < data[0].Count; i++)
                {
                    dataGridView1.Rows.Add(data[i].ToArray());
                }
            }*/
        }

        private void button8_Click(object sender, EventArgs e)
        {
            db.selectall("Structure_Academic_Discipline", dataGridView1);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //тут трабл якщо натискаєш на пусту клітинку
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

            db.insert("Structure_Academic_Discipline", fields, values);
            db.selectall("Structure_Academic_Discipline", dataGridView1);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            db.delete("Structure_Academic_Discipline", "Num_Section=" + textBox1.Text);
            db.selectall("Structure_Academic_Discipline", dataGridView1);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            String[] fields = { "Num_Section", "Num_Class", "Name_Section", "Total_Hours", "Lecture_Hours", "Workshop_Hours",
             "Practical_Hours","Laboratory_Hours", "IndepWorkStud_Hours", "Recommended_Books", "Forms_Means_Con" };
            String[] values = { textBox1.Text, textBox2.Text, "'" + textBox3.Text + "'", textBox4.Text,
            textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text,
             "'" + textBox10.Text + "'", "'" + textBox11.Text + "'"};

            db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text);
            db.selectall("Structure_Academic_Discipline", dataGridView1);
        }
    }
}