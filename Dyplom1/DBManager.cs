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
        String conStr = @"URI=file:D:\\sqllite\\Diplom\\Diplom.db";
        public DBManager()
        {
            connection = new SQLiteConnection();
            command = new SQLiteCommand();
            command.Connection = connection;
        }
        public void connectTo(String pconStr)
        {
            connection.ConnectionString = pconStr;
        }

        public void connectTo()
        {
            connection.ConnectionString = conStr;   
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
                /*
                Console.WriteLine("it's working, for sure");
                var res = new List<List<object>>();
                connection.Open();
                command.CommandText = "select *  from " + tablename;
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    List<Object> row = new List<object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        row.Add(reader[i]);
                    }
                    res.Add(row);
                };
                connection.Close();
                return res;*/
                
                command.CommandText = "SELECT Num_Section \"№ Розділу\", Num_Class \"№ Заняття\", Name_Section \"Назви розділів\", Total_Hours \"Усього годин\", Lecture_Hours \"Лекційні години\", Workshop_Hours \"Семінарські години\", Practical_Hours \"Практичні години\", Laboratory_Hours \"Лабораторні години\", IndepWorkStud_Hours \"С.р.с години\", Recommended_Books \"Рекомендована література\", Forms_Means_Con \"Форми та засоби контролю\" FROM " + tablename;
                MessageBox.Show(command.CommandText);
                connection.Open();
                SQLiteDataReader datareader = command.ExecuteReader();
                fillgrid(datareader, datagrid);
                datareader.Close();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
                //тут та ж помилка з датарідеором, що поки він активний то не можна встановити commandtext
            }
        }
        public List<List<Object>> selectall(String tablename)
        {
            try
            {
                var res = new List<List<Object>>();
                command.CommandText = "SELECT Num_Section \"№ Розділу\", Num_Class \"№ Заняття\", Name_Section \"Назви розділів\", Total_Hours \"Усього годин\", Lecture_Hours \"Лекційні години\", Workshop_Hours \"Семінарські години\", Practical_Hours \"Практичні години\", Laboratory_Hours \"Лабораторні години\", IndepWorkStud_Hours \"С.р.с години\", Recommended_Books \"Рекомендована література\", Forms_Means_Con \"Форми та засоби контролю\" FROM " + tablename;
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
                MessageBox.Show(command.CommandText);
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
                MessageBox.Show(command.CommandText);
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
                MessageBox.Show(command.CommandText);
                command.ExecuteNonQuery();
                connection.Close();
               // }
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
                MessageBox.Show(command.CommandText);
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string sum(String tablename, String[] fields, String[] values, String cond)
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
                
                  string query = "SELECT SUM(Total_Hours) AS TotalHours FROM " + tablename;
                if (!string.IsNullOrEmpty(cond))
                {
                    query += " WHERE " + cond;
                }

                command.CommandText = query;
                string result = command.ExecuteScalar()?.ToString();
                connection.Close();

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
                 
               /* if (cond != null)
                {
                command.CommandText = "select sum(Total_Hours) Усього годин from " + tablename + " where " + cond;
                }
                else
                {
                command.CommandText = "select sum(" + String.Join(",", nonEmptyFields) + ") from " + tablename;
                    tmp = String.Join(",", nonEmptyFields);
                }
                MessageBox.Show(command.CommandText);
                command.ExecuteNonQuery();
                connection.Close();
                return result;
            }
            catch(Exception ex)
            {
                throw ex;
            }*/
        }
        public string _getindex(String tablename, string section_, string class_)
        {
            //ось так?
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
            //SELECT "Index" FROM Structure_Academic_Discipline INDEXED BY Index_Section_Class WHERE Num_Section = 3 AND Num_Class = 22;
            return index;
        }
    }
}