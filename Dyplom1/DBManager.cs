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
                command.CommandText = "select * from " + tablename;
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
                command.CommandText = "select * from " + tablename;
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
               // SQLiteDataReader datareader = command.ExecuteReader();
               // datareader.Close();

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
                //System.InvalidOperationException: 'Cannot set CommandText while a DataReader is active'
                //видає логічну помилку десь біля коми (стара помилка)
                //тепер видає іншу помилку: Cannot set CommandText while a DataReader is active

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
        public void update(String tablename, String[] fields, String[] values, String cond)
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
                    + " = " + nonEmptyValues[nonEmptyValues.Count - 1] + " where " + cond;
               
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}