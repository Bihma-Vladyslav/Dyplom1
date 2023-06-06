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
        int index = 0;
        string ind = "";
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
            if(dataGridView1.Rows.Count != e.RowIndex + 1)
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
                    ind = db._getindex("Structure_Academic_Discipline", textBox1.Text, textBox2.Text);
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
            db.delete("Structure_Academic_Discipline", "Num_Class=" + textBox2.Text);
            db.selectall("Structure_Academic_Discipline", dataGridView1);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            String[] fields = { "Num_Section", "Num_Class", "Name_Section", "Total_Hours", "Lecture_Hours", "Workshop_Hours",
             "Practical_Hours","Laboratory_Hours", "IndepWorkStud_Hours", "Recommended_Books", "Forms_Means_Con" };
            String[] values = { textBox1.Text, textBox2.Text, "'" + textBox3.Text + "'", textBox4.Text,
            textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text,
             "'" + textBox10.Text + "'", "'" + textBox11.Text + "'"};
            db.update("Structure_Academic_Discipline", fields, values, "\"Index\"" , ind);
            db.selectall("Structure_Academic_Discipline", dataGridView1);      
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
            listBox2.Items.Clear();
            index = 1;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string [] sumResult = db.Sum("Structure_Academic_Discipline", fields, values, condition);

            listBox2.Items.Add("Разом за розділом (змістовим модулем) 1");
            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            index = 2;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.Sum("Structure_Academic_Discipline", fields, values, condition);

            listBox2.Items.Add("Разом за розділом (змістовим модулем) 2");
            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            index = 3;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.Sum("Structure_Academic_Discipline", fields, values, condition);

            listBox2.Items.Add("Разом за розділом (змістовим модулем) 3");
            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            index = 4;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.Sum("Structure_Academic_Discipline", fields, values, condition);

            listBox2.Items.Add("Разом за розділом (змістовим модулем) 4");
            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            index = 5;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.Sum("Structure_Academic_Discipline", fields, values, condition);

            listBox2.Items.Add("Разом за розділом (змістовим модулем) 5");
            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            index = 6;
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.Sum("Structure_Academic_Discipline", fields, values, condition);

            listBox2.Items.Add("Разом за розділом (змістовим модулем) 6");
            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours", "Practical_Hours", "Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = new string[7];

            string condition = "Num_Section = " + index;
            string[] sumResult = db.SumAll("Structure_Academic_Discipline", fields, values);

            listBox2.Items.Add("Усього годин: " + sumResult[0]);
            listBox2.Items.Add("Лекційні години: " + sumResult[1]);
            listBox2.Items.Add("Семінарські години: " + sumResult[2]);
            listBox2.Items.Add("Практичні години: " + sumResult[3]);
            listBox2.Items.Add("Лабораторні години: " + sumResult[4]);
            listBox2.Items.Add("С.р.с години: " + sumResult[5]);
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
    }
}