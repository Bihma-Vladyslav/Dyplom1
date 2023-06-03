using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SQLite;

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

        public List<List<Object>> Selectall(String tablename)
        {
            try
            {
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
                return res;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}