using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Data;

namespace парик
{
    internal class DB
    {
        public static MySqlConnection connection = new MySqlConnection(@"Server = localhost; dataBase= parik; port=3306; User id=root; password = Loowe");
        public void openConnetion()
        {
            if (connection.State == System.Data.ConnectionState.Open)
            {
                connection.Open();
            }
        }
        public void closeConnetion()
        {
            if (connection.State == System.Data.ConnectionState.Closed)
            {
                connection.Close();
            }
        }
        public MySqlConnection GetMySqlConnection()
        {
            return connection;
        }
    }
}
