using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DataTable = System.Data.DataTable;
using Window = System.Windows.Window;

namespace парик
{
    /// <summary>
    /// Логика взаимодействия для Чек.xaml
    /// </summary>
    public partial class Чек : Window
    {
        public Чек()
        {
            InitializeComponent();
            daaa.SelectedDate = DateTime.Today;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Loadch();
        }
        private void Loadch()
        {
            MySqlDataAdapter adapter = new MySqlDataAdapter("select idKlienta as 'Номер', FIO as 'ФИО', Telefon as 'Телефон', kolvo as 'Количесвто посещений' from klients", DB.connection);
            DataTable tables = new DataTable();
            adapter.Fill(tables);
            kl.ItemsSource = tables.DefaultView;

            MySqlDataAdapter adapter1 = new MySqlDataAdapter("select idCheka as 'Номер чека', idSotrudnika as 'Номер сотрудника', idKlienta as 'Номер клиента', datas as 'Дата' from chek", DB.connection);
            DataTable tables1 = new DataTable();
            adapter1.Fill(tables1);
            chek.ItemsSource = tables1.DefaultView;

            kolvv.Visibility = Visibility.Hidden;
        }

        private void kl_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)kl.SelectedItems[0];

                DB.connection.Open();
                System.Data.DataTable table = new System.Data.DataTable();
                MySqlDataAdapter adapter = new MySqlDataAdapter($"select idKlienta,FIO, kolvo from klients where idKlienta = \"{Convert.ToInt32(row["Номер"])}\"", DB.connection);
                adapter.Fill(table);
                fioId.Text = table.Rows[0][0].ToString();
                fiokl.Text = table.Rows[0][1].ToString();
                kolvv.Text = table.Rows[0][2].ToString();

                DB.connection.Close();
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult messageBoxResult = MessageBox.Show("Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (messageBoxResult == MessageBoxResult.Yes)
            {
                admin admin = new admin();
                admin.Show();
                Hide();
            }
        }
        string data = DateTime.Today.ToString("yyyy-MM-dd");

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row1 = (DataRowView)kl.SelectedItems[0];
                DB.connection.Open();
                MySqlCommand update = new MySqlCommand("update Klients set kolvo = @kolvo + 1 where idKlienta = @idKlienta", DB.connection);
                update.Parameters.AddWithValue("idKlienta", Convert.ToInt32(row1["Номер"]));
                update.Parameters.AddWithValue("kolvo", kolvv.Text);
                update.ExecuteNonQuery();
                DB.connection.Close();
                Loadch();
            }
            catch
            {
                MessageBox.Show("Выберите элемент из таблицы и нажмите на кнопку повторно", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }

            int a = Convert.ToInt32(kolvv.Text);
            if (a > 4)
            {
                MessageBox.Show("Клиент приходил больше 5 раз и он получает скидку в 5%", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                DB.connection.Open();
                string da = $"insert into Chek(idSotrudnika, idKlienta, datas) values (@idSotrudnika, @idKlienta, @datas)";
                MySqlCommand insert = new MySqlCommand(da, DB.connection);
                insert.Parameters.AddWithValue("idKlienta", fioId.Text);
                insert.Parameters.AddWithValue("idSotrudnika", ID.idSotrudnika);
                insert.Parameters.AddWithValue("datas", DateTime.Now.ToString("yyy-MM-dd"));
                insert.ExecuteNonQuery();
                Loadch();

                DataRowView row = (DataRowView)kl.Items[kl.Items.Count - 1];
                DataRowView row2 = (DataRowView)chek.Items[chek.Items.Count - 1];

                int idDogovor = Convert.ToInt32(row2[0].ToString());

                idCheka.idChek = idDogovor;

                Выбрать_услугу выбрать_Услугу = new Выбрать_услугу(row[0].ToString());
                выбрать_Услугу.Owner = this;
                выбрать_Услугу.Show();

                Hide();
                DB.connection.Close();
            }
            else
            {
                DB.connection.Open();
                string da = $"insert into Chek(idSotrudnika, idKlienta, datas) values (@idSotrudnika, @idKlienta, @datas)";
                MySqlCommand insert = new MySqlCommand(da, DB.connection);
                insert.Parameters.AddWithValue("idKlienta", fioId.Text);
                insert.Parameters.AddWithValue("idSotrudnika", ID.idSotrudnika);
                insert.Parameters.AddWithValue("datas", DateTime.Now.ToString("yyy-MM-dd"));
                insert.ExecuteNonQuery();
                Loadch();

                DataRowView row = (DataRowView)kl.Items[kl.Items.Count - 1];
                DataRowView row2 = (DataRowView)chek.Items[chek.Items.Count - 1];

                int idDogovor = Convert.ToInt32(row2[0].ToString());

                idCheka.idChek = idDogovor;

                выб выбрать_Услугу = new выб(row[0].ToString());
                выбрать_Услугу.Owner = this;
                выбрать_Услугу.Show();

                Hide();
                DB.connection.Close();
            }
        }
    }
}
