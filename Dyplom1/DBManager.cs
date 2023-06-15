using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SQLite;
using System.Windows.Forms;

namespace Dyplom1
{
    
    public class DBManager
    {
        SQLiteConnection connection;
        SQLiteCommand command;
        String conStr;
        public DBManager()
        {
            conStr = @"URI=file:" + System.AppDomain.CurrentDomain.BaseDirectory + "Diplom.db";
            connection = new SQLiteConnection();
            command = new SQLiteCommand();
            command.Connection = connection;
            if(!CheckIfTableExists(conStr,"Structure_Academic_Discipline"))
            {
                CreateTable(conStr);
            }
        }
        public void connectTo(String pconStr)
        {
            connection.ConnectionString = pconStr;
        }

        public void connectTo()
        {
            connection.ConnectionString = conStr;   
        }

        static bool CheckIfTableExists(string connectionString, string tableName)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                SQLiteCommand command = new SQLiteCommand();
                command.Connection = connection;
                command.CommandText = $"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName}'";

                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    return reader.HasRows;
                }
            }
        }
        static void CreateTable(string connectionString)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();

                SQLiteCommand command = new SQLiteCommand();
                command.Connection = connection;
                command.CommandText = $"CREATE TABLE \"Structure_Academic_Discipline\" ( \"Num_Section\"	INTEGER NOT NULL, \"Num_Class\"	INTEGER NOT NULL, \"Name_Section\"	TEXT NOT NULL, \"Total_Hours\"	NUMERIC NOT NULL, \"Lecture_Hours\"	INTEGER, \"Workshop_Hours\"	INTEGER, \"Practical_Hours\"	INTEGER, \"Laboratory_Hours\"	TEXT, \"IndepWorkStud_Hours\"	TEXT, \"Recommended_Books\"	TEXT, \"Forms_Means_Con\"	TEXT, \"Index\"	INTEGER NOT NULL, PRIMARY KEY(\"Index\" AUTOINCREMENT) );";

                command.ExecuteNonQuery();

                command.CommandText = $"CREATE TABLE \"Topics_Seminar_Classes\" ( \"Number_Sequence\"	INTEGER NOT NULL UNIQUE, \"Topic_Name\"	TEXT, \"Number_Hours\"	INTEGER, PRIMARY KEY(\"Number_Sequence\") );";

                command.ExecuteNonQuery();

                command.CommandText = $"CREATE TABLE \"Topics_Practical_Classes\" ( \"Number_Sequence\"	INTEGER NOT NULL UNIQUE, \"Topic_Name\"	TEXT, \"Number_Hours\"	INTEGER, PRIMARY KEY(\"Number_Sequence\") );";

                command.ExecuteNonQuery();

                command.CommandText = $"CREATE TABLE \"Topics_Laboratory_Works\" ( \"Number_Sequence\"	INTEGER NOT NULL UNIQUE, \"Name_Class\"	TEXT NOT NULL, \"Number_Hours\"	INTEGER NOT NULL, PRIMARY KEY(\"Number_Sequence\") );";

                command.ExecuteNonQuery();

                command.CommandText = $"CREATE TABLE \"Topics_Independent_Works\" ( \"Number_Sequence\"	INTEGER NOT NULL UNIQUE, \"Name_Class\"	TEXT NOT NULL, \"Number_Hours\"	INTEGER NOT NULL, PRIMARY KEY(\"Number_Sequence\") );";

                command.ExecuteNonQuery();

                command.CommandText = $"CREATE UNIQUE INDEX \"Index_Section_Class\" ON \"Structure_Academic_Discipline\" ( \"Num_Section\", \"Num_Class\" );";

                command.ExecuteNonQuery();
            }
        }
        

        public void fillgrid(SQLiteDataReader datareader, DataGridView datagrid)
        {
            datagrid.Columns.Clear();
            for (int i = 0; i < datareader.FieldCount; i++)
            {
                datagrid.Columns.Add("col" + i.ToString(), datareader.GetName(i));
            }
            while (datareader.Read())
            {
                String[] s = new String[datareader.FieldCount];
                for (int i = 0; i < datareader.FieldCount; i++)
                {
                    s[i] = datareader[i].ToString();
                }
                datagrid.Rows.Add(s);
            };
        }
        public void selectall(String tablename, DataGridView datagrid)
        {
            try
            {               
                command.CommandText = "SELECT Num_Section \"№ Розділу\", Num_Class \"№ Заняття\", Name_Section \"Назви розділів\", " +
                    "Total_Hours \"Усього годин\", Lecture_Hours \"Лекційні години\", Workshop_Hours \"Семінарські години\", " +
                    "Practical_Hours \"Практичні години\", Laboratory_Hours \"Лабораторні години\", IndepWorkStud_Hours \"С.р.с години\", " +
                    "Recommended_Books \"Рекомендована література\", Forms_Means_Con \"Форми та засоби контролю\" FROM " + tablename;
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                fillgrid(datareader, datagrid);
                datareader.Close();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<List<Object>> selectall(String tablename)
        {
            try
            {
                var res = new List<List<Object>>();
                command.CommandText = "SELECT Num_Section \"№ Розділу\", Num_Class \"№ Заняття\", Name_Section \"Назви розділів\", " +
                    "Total_Hours \"Усього годин\", Lecture_Hours \"Лекційні години\", Workshop_Hours \"Семінарські години\"," +
                    " Practical_Hours \"Практичні години\", Laboratory_Hours \"Лабораторні години\", IndepWorkStud_Hours \"С.р.с години\", " +
                    "Recommended_Books \"Рекомендована література\", Forms_Means_Con \"Форми та засоби контролю\" FROM " + tablename;
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                while (datareader.Read())
                {
                    List<Object> row = new List<object>();
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        row.Add(datareader[i]);
                    }
                    res.Add(row);
                }
                datareader.Close();
                connection.Close();
                return res;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void selectall1(String tablename, DataGridView datagrid)
        {
            try
            {
                command.CommandText = "SELECT Number_Sequence \"№ з/п\", Topic_Name \"Назва теми\", Number_Hours \"Кількість годин\" FROM " + tablename;
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                fillgrid(datareader, datagrid);
                datareader.Close();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<List<Object>> selectall1(String tablename)
        {
            try
            {
                var res = new List<List<Object>>();
                command.CommandText = "SELECT Number_Sequence \"№ з/п\", Topic_Name \"Назва теми\", Number_Hours \"Кількість годин\" FROM " + tablename;
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                while (datareader.Read())
                {
                    List<Object> row = new List<object>();
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        row.Add(datareader[i]);
                    }
                    res.Add(row);
                }
                datareader.Close();
                connection.Close();
                return res;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void selectall2(String tablename, DataGridView datagrid)
        {
            try
            {
                command.CommandText = "SELECT Number_Sequence \"пн\", Name_Class \"Назва заняття\", Number_Hours \"Кількість годин\" FROM " + tablename;
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                fillgrid(datareader, datagrid);
                datareader.Close();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public List<List<Object>> selectall2(String tablename)
        {
            try
            {
                var res = new List<List<Object>>();
                command.CommandText = "SELECT Number_Sequence \"пн\", Name_Class \"Назва заняття\", Number_Hours \"Кількість годин\" FROM " + tablename;
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                while (datareader.Read())
                {
                    List<Object> row = new List<object>();
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        row.Add(datareader[i]);
                    }
                    res.Add(row);
                }
                datareader.Close();
                connection.Close();
                return res;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
   
        public void ExecSQl(String query)
        {
            try
            {
                command.CommandText = query;
                connection.Open();
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void insert(String tablename, String[] fields, String[] values)
        {
            try
            {
                connection.Open();

                List<string> nonEmptyFields = new List<string>();
                List<string> nonEmptyValues = new List<string>();

                for (int i = 0; i < fields.Length; i++)
                {
                    if (!string.IsNullOrEmpty(values[i]))
                    {
                        nonEmptyFields.Add(fields[i]);
                        nonEmptyValues.Add(values[i]);
                    }
                }

                command.CommandText = "insert into "+tablename+"("+String.Join(",", nonEmptyFields) +") values("+ 
                    String.Join(",", nonEmptyValues)+")";
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void delete(String tablename, String cond)
        {
            try
            {
                connection.Open();
                if (cond != null)
                {
                    command.CommandText = "delete from " + tablename + " where " + cond;
                }
                else
                {
                    command.CommandText = "delete from " + tablename;
                }
                
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка! Будь ласка, обов'язково заповніть поле «Назва розділу»" +
                    "для видалення необіхдного поля");
                throw ex;
            }
        }
        public void update(String tablename, String[] fields, String[] values, String column, String value)
        {
            try
            {
                connection.Open();

                List<string> nonEmptyFields = new List<string>();
                List<string> nonEmptyValues = new List<string>();

                for (int i = 0; i < fields.Length; i++)
                {
                    if (!string.IsNullOrEmpty(values[i]))
                    {
                        nonEmptyFields.Add(fields[i]);
                        nonEmptyValues.Add(values[i]);
                    }
                }
                command.CommandText = "update " + tablename + " set ";
                for (int i = 0; i < nonEmptyFields.Count - 1 && i < nonEmptyValues.Count - 1; i++)
                {
                    command.CommandText += nonEmptyFields[i] + " = " + nonEmptyValues[i] + " , ";
                }
                command.CommandText += nonEmptyFields[nonEmptyFields.Count - 1]
                    + " = " + nonEmptyValues[nonEmptyValues.Count - 1] + " where" + column + " = " + value;;
                string tmp = command.CommandText;
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string[] sum(string tableName, string[] fields, string[] values, string condition)
        {
            try
            {
                connection.Open();

                List<string> nonEmptyFields = new List<string>();
                List<string> nonEmptyValues = new List<string>();

                for (int i = 0; i < fields.Length; i++)
                {
                    if (!string.IsNullOrEmpty(values[i]))
                    {
                        nonEmptyFields.Add(fields[i]);
                        nonEmptyValues.Add(values[i]);
                    }
                }

                string query = "SELECT SUM(Total_Hours) \"Усього годин\", SUM(Lecture_Hours) \"Лекційні години\", SUM(Workshop_Hours) \"Семінарські години\"," +
                    " SUM(Practical_Hours) \"Практичні години\", SUM(Laboratory_Hours) \"Лабораторні години\", SUM(IndepWorkStud_Hours) \"С.р.с години\" FROM " + tableName;
                if (!string.IsNullOrEmpty(condition))
                {
                    query += " WHERE " + condition;
                }

                command.CommandText = query;
                string[] result = null;
                SQLiteDataReader datareader = command.ExecuteReader();
                while (datareader.Read())
                {
                    result = new String[datareader.FieldCount];
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        result[i] = datareader[i].ToString();
                    }
                };
                datareader.Close();
                connection.Close();
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string[] sumall(string tableName, string[] fields, string[] values)
        {
            try
            {
                connection.Open();

                List<string> nonEmptyFields = new List<string>();
                List<string> nonEmptyValues = new List<string>();

                for (int i = 0; i < fields.Length; i++)
                {
                    if (!string.IsNullOrEmpty(values[i]))
                    {
                        nonEmptyFields.Add(fields[i]);
                        nonEmptyValues.Add(values[i]);
                    }
                }

                string query = "SELECT SUM(Total_Hours) \"Усього годин\", SUM(Lecture_Hours) \"Лекційні години\", SUM(Workshop_Hours) \"Семінарські години\"," +
                    " SUM(Practical_Hours) \"Практичні години\", SUM(Laboratory_Hours) \"Лабораторні години\", SUM(IndepWorkStud_Hours) \"С.р.с години\" FROM " + tableName;

                command.CommandText = query;
                string[] result = null;
                SQLiteDataReader datareader = command.ExecuteReader();
                while (datareader.Read())
                {
                    result = new String[datareader.FieldCount];
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        result[i] = datareader[i].ToString();
                    }
                };
                datareader.Close();
                connection.Close();
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string getindex(String tablename, string section_, string class_)
        {
            string index = "";
            try
            {
                
                command.CommandText = "select \"Index\" FROM Structure_Academic_Discipline INDEXED BY Index_Section_Class WHERE Num_Section = " + section_ + " AND Num_Class = " + class_;
                connection.Open();

                SQLiteDataReader dr = command.ExecuteReader();
                List<Object> value = new List<Object>(); 
                while(dr.Read())
                {
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        value.Add(dr[i]);
                    }
                }

                index = value[0].ToString();
                dr.Close();
                connection.Close();

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return index;
        }
    }
}