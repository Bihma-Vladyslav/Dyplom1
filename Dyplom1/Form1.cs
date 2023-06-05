﻿using System;
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
            db.update("Structure_Academic_Discipline", fields, values, ind);
            //, "Num_Section=" + textBox1.Text +             " and Name_Section=" + "'" + textBox3.Text + "'"
            db.selectall("Structure_Academic_Discipline", dataGridView1);
            /*
            if (textBox1.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + 
                    " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox2.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox3.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox4.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox5.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox6.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox7.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox8.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox9.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox10.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");

            }

            if (textBox11.Tag.ToString() == "1")
            {
                db.update("Structure_Academic_Discipline", fields, values, "Num_Section=" + textBox1.Text + " and Name_Section=" + "'" + textBox3.Text + "'");
                db.selectall("Structure_Academic_Discipline", dataGridView1);
            }
            
            textBox1.Tag = 0;
            textBox2.Tag = 0;
            textBox3.Tag = 0;
            textBox4.Tag = 0;
            textBox5.Tag = 0;
            textBox6.Tag = 0;
            textBox7.Tag = 0;
            textBox8.Tag = 0;
            textBox9.Tag = 0;
            textBox10.Tag = 0;
            textBox11.Tag = 0;
            */
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
            String[] fields = { "Total_Hours", "Lecture_Hours", "Workshop_Hours",
             "Practical_Hours","Laboratory_Hours", "IndepWorkStud_Hours" };
            String[] values = { textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text};
            db.sum("Structure_Academic_Discipline", fields, values, "Num_Section=" + index);
            listBox2.Items.Add("Разом за розділом (змістовим модулем) 1");
            listBox2.Items.Add("Total_Hours: " + textBox4.Text);
            listBox2.Items.Add("Lecture_Hours: " + textBox5.Text);
            listBox2.Items.Add("Workshop_Hours: " + textBox6.Text);
            listBox2.Items.Add("Practical_Hours: " + textBox7.Text);
            listBox2.Items.Add("Laboratory_Hours: " + textBox8.Text);
            listBox2.Items.Add("IndepWorkStud_Hours: " + textBox9.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("Разом за розділом (змістовим модулем) 2");
            listBox2.Items.Add("Total_Hours: ");
            listBox2.Items.Add("Lecture_Hours: ");
            listBox2.Items.Add("Workshop_Hours: ");
            listBox2.Items.Add("Practical_Hours: ");
            listBox2.Items.Add("Laboratory_Hours: ");
            listBox2.Items.Add("IndepWorkStud_Hours: ");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("Разом за розділом (змістовим модулем) 3");
            listBox2.Items.Add("Total_Hours: ");
            listBox2.Items.Add("Lecture_Hours: ");
            listBox2.Items.Add("Workshop_Hours: ");
            listBox2.Items.Add("Practical_Hours: ");
            listBox2.Items.Add("Laboratory_Hours: ");
            listBox2.Items.Add("IndepWorkStud_Hours: ");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("Разом за розділом (змістовим модулем) 4");
            listBox2.Items.Add("Total_Hours: ");
            listBox2.Items.Add("Lecture_Hours: ");
            listBox2.Items.Add("Workshop_Hours: ");
            listBox2.Items.Add("Practical_Hours: ");
            listBox2.Items.Add("Laboratory_Hours: ");
            listBox2.Items.Add("IndepWorkStud_Hours: ");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("Разом за розділом (змістовим модулем) 5");
            listBox2.Items.Add("Total_Hours: ");
            listBox2.Items.Add("Lecture_Hours: ");
            listBox2.Items.Add("Workshop_Hours: ");
            listBox2.Items.Add("Practical_Hours: ");
            listBox2.Items.Add("Laboratory_Hours: ");
            listBox2.Items.Add("IndepWorkStud_Hours: ");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("Разом за розділом (змістовим модулем) 6");
            listBox2.Items.Add("Total_Hours: ");
            listBox2.Items.Add("Lecture_Hours: ");
            listBox2.Items.Add("Workshop_Hours: ");
            listBox2.Items.Add("Practical_Hours: ");
            listBox2.Items.Add("Laboratory_Hours: ");
            listBox2.Items.Add("IndepWorkStud_Hours: ");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listBox2.Items.Add("Усього годин");
            listBox2.Items.Add("Total_Hours: ");
            listBox2.Items.Add("Lecture_Hours: ");
            listBox2.Items.Add("Workshop_Hours: ");
            listBox2.Items.Add("Practical_Hours: ");
            listBox2.Items.Add("Laboratory_Hours: ");
            listBox2.Items.Add("IndepWorkStud_Hours: ");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {

         MessageBox.Show(db._getindex("Structure_Academic_Discipline", "3","22").ToString());
        }
    }
}