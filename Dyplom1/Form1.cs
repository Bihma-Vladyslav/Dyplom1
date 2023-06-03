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
            var data = db.Selectall("Structure_Academic_Discipline");
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
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var data = db.Selectall("Structure_Academic_Discipline");
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
            }
        }
    }
}