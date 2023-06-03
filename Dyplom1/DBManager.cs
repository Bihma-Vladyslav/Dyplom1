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
                command.CommandText = "insert into "+tablename+"("+String.Join(",", fields) +") values("+ 
                    String.Join(",",values)+")";
                connection.Open();
                command.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
                //видає логічну помилку десь біля коми (стара помилка)
                //тепер видає іншу помилку: Cannot set CommandText while a DataReader is active

                throw ex;
            }
        }
        public void delete(String tablename, String cond)
        {
            try
            {
                if (cond != null)
                {
                    command.CommandText = "delete from " + tablename + " where " + cond;
                }
                else
                {
                    command.CommandText = "delete from " + tablename;
                }
                connection.Open();
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
                command.CommandText = "update " + tablename + " set ";
                for (int i = 0; i < fields.Length - 1 && i < values.Length - 1; i++)
                {
                    command.CommandText += fields[i] + " = " + values[i] + " , ";
                }
                command.CommandText += fields[fields.Length - 1] + " = " + values[values.Length - 1] + " where " + cond;
                connection.Open();
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