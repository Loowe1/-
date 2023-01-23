using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace парик
{
    /// <summary>
    /// Логика взаимодействия для выб.xaml
    /// </summary>
    public partial class выб : Window
    {
        public выб(string UserID)
        {
            InitializeComponent();
            DB db = new DB();
            UseridBox.Text = Convert.ToString(idCheka.idChek);
            fiosot.Text = ID.FIO;
        }

        private void uss_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)uss.SelectedItems[0];

                DB.connection.Open();
                System.Data.DataTable table = new System.Data.DataTable();
                MySqlDataAdapter ad1 = new MySqlDataAdapter($"select IDusugi, Stoimost from Usugi where IDusugi = \"{Convert.ToInt32(row["Номер"])}\"", DB.connection);
                ad1.Fill(table);
                nomus.Text = table.Rows[0][0].ToString();
                sto.Text = table.Rows[0][1].ToString();

                DB.connection.Close();
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void ch_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            DataRowView row = (DataRowView)ch.SelectedItems[0];

            DB.connection.Open();
            System.Data.DataTable table = new System.Data.DataTable();
            MySqlDataAdapter ad1 = new MySqlDataAdapter($"select idTabl as 'Номер табличной части', usugi.IDusugi as 'Номер услуги', usugi.Nazvanie as 'Название услуги', usugi.Opisanie as 'Описание', usugi.Stoimost as 'Стоимость', tabl.idCheka as 'Номер чека' from tabl inner join usugi on tabl.IDusugi = usugi.IDusugi where tabl.idCheka = '{UseridBox.Text}'", DB.connection);
            ad1.Fill(table);

            DB.connection.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Loadvi();
            Чек abc = this.Owner as Чек;
            if (abc != null) datt.Text = abc.daaa.Text;
            if (abc != null) fiokl1.Text = abc.fiokl.Text;
        }
        private void Loadvi()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter("select IDusugi as 'Номер', Nazvanie as 'Название', Stoimost as 'Стоимость, руб', Opisanie as 'Описание' from usugi order by IDusugi ", DB.connection);
            System.Data.DataTable tables = new System.Data.DataTable();
            adapter.Fill(tables);
            uss.ItemsSource = tables.DefaultView;

            UseridBox.Visibility = Visibility.Hidden;
            nomus.Visibility = Visibility.Hidden;
            fiosot.Visibility = Visibility.Hidden;
            sto.Visibility = Visibility.Hidden;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (UseridBox.Text == "" || fiosot.Text == "" || fiokl1.Text == "" || datt.Text == "" || nomus.Text == "")
            {
                MessageBox.Show("Все поля должны быть заполнены!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                DB.connection.Open();

                string sql = $"insert into tabl (IDusugi, idCheka) values (@IDusugi, @idCheka)";
                MySqlCommand insert = new MySqlCommand(sql, DB.connection);

                insert.Parameters.AddWithValue("IDusugi", nomus.Text);
                insert.Parameters.AddWithValue("idCheka", UseridBox.Text);
                insert.ExecuteNonQuery();

                MySqlDataAdapter adapter = new MySqlDataAdapter($"select idTabl as 'Номер табличной части', usugi.IDusugi as 'Номер услуги', usugi.Nazvanie as 'Название услуги', usugi.Opisanie as 'Описание', usugi.Stoimost as 'Стоимость', tabl.idCheka as 'Номер чека' from tabl inner join usugi on tabl.IDusugi = usugi.IDusugi where tabl.idCheka = '{UseridBox.Text}'", DB.connection);
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
                ch.ItemsSource = dataTable.DefaultView;

                DB.connection.Close();
                Loadvi();

                nomus.Clear();
            }
        }

        private void Button_Click3(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (messageBoxResult == MessageBoxResult.Yes)
            {
                Чек admin = new Чек();
                admin.Show();
                Hide();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var Excel = new Microsoft.Office.Interop.Excel.Application();
            var abc = Excel.Workbooks.Open(@"C:\Users\shaki\Desktop\My Fucking Works\парик\чек без скидки.xlsx");
            var rew = abc.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)rew.Cells[28, 4];
            Microsoft.Office.Interop.Excel.Range range1 = range.EntireRow;
            range1.Select();

            Microsoft.Office.Interop.Excel.Range range2 = (Microsoft.Office.Interop.Excel.Range)rew.Cells[35, 5];
            Microsoft.Office.Interop.Excel.Range range3 = range.EntireRow;

            System.Data.DataTable dataTable = new System.Data.DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"select idTabl as 'Номер табличной части', usugi.IDusugi as 'Номер услуги', usugi.Nazvanie as 'Название услуги', usugi.Opisanie as 'Описание', usugi.Stoimost as 'Стоимость', tabl.idCheka as 'Номер чека' from tabl inner join usugi on tabl.IDusugi = usugi.IDusugi where tabl.idCheka = '{UseridBox.Text}'", DB.connection);
            adapter.Fill(dataTable);

            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                range3.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, Missing.Value); //вставляет пустую такую же строку
            }
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                rew.Cells[28 + i, 4] = dataTable.Rows[i][2];
                rew.Cells[28 + i, 5] = dataTable.Rows[i][4];
            }
            rew.Cells[18, 5] = UseridBox.Text;
            rew.Cells[21, 4] = fiokl1.Text;
            rew.Cells[25, 4] = datt.Text.Replace("0:00:00", " ");
            rew.Cells[23, 4] = fiosot.Text;
            int ass = 28;
            int ass1 = 28;

            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                ass1 += 1;
                rew.Cells[35 + i, 5].FormulaLocal = $"=СУММ(E{ass}:E{ass1})";
            }
            Excel.Visible = true;
        }
    }
}
